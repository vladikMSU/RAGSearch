from __future__ import annotations

import argparse
import datetime as dt
import ipaddress
import json
import os
import secrets
import shutil
import stat
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, IO, Mapping
from urllib.parse import urlsplit

from ragsearch_service.config import Settings


DEFAULT_EXECUTABLE = (
    Path(__file__).resolve().parent.parent
    / "native-mapi-probe"
    / "build-direct"
    / "NativeMapiProbe.exe"
)
DEFAULT_SERVICE_URL = "http://127.0.0.1:8765"
TOKEN_HEADER = "X-RAGSearch-Token"
PROGRESS_PREFIX = "RAGSEARCH_PROGRESS "

DEFAULT_MAX_STORES = 16
DEFAULT_MAX_FOLDERS = 100
DEFAULT_MAX_MESSAGES = 1_000
DEFAULT_BODY_CHARS = 200_000

MAX_STORES = 128
MAX_FOLDERS = 100_000
MAX_MESSAGES = 100_000
MAX_BODY_CHARS = 4_000_000
MAX_STORE_FILTER_CHARS = 256
MAX_HTTP_RESPONSE_BYTES = 1_048_576
MAX_HTTP_REQUEST_BYTES = Settings.default().request_limit_bytes
HTTP_TIMEOUT_SECONDS = 300

MAX_ID_CHARS = 32_768
MAX_FOLDER_PATH_CHARS = 32_768
MAX_STORE_NAME_CHARS = 4_096
MAX_SUBJECT_CHARS = 1_000_000
MAX_ADDRESS_FIELD_CHARS = 1_000_000
MAX_METADATA_FIELD_CHARS = 65_536
MAX_TIMESTAMP_CHARS = 64
MAX_ATTACHMENTS_PER_MESSAGE = 4_096
MAX_ATTACHMENT_NAME_CHARS = 32_768
MAX_CONTENT_TYPE_CHARS = 4_096
MAX_TEMP_PATH_CHARS = 32_768
MAX_DECLARED_ATTACHMENT_BYTES = (1 << 63) - 1
MAX_JSONL_OVERHEAD_CHARS = 8 * 1024 * 1024

OWNED_RUN_PREFIX = "native-import-"
OWNERSHIP_MARKER_NAME = ".ragsearch-native-import-owned"


class ImportFailure(RuntimeError):
    """A safe, user-facing native import failure."""


class ImportCancelled(RuntimeError):
    """The caller requested a graceful native import cancellation."""


@dataclass(frozen=True)
class ImportResult:
    imported: int
    probe_exit_code: int
    bodies_truncated: int = 0


@dataclass(frozen=True)
class OwnedSpoolRun:
    base_dir: Path
    run_dir: Path
    marker_token: str


ProgressCallback = Callable[[Mapping[str, object]], None]


def _bounded_integer(option: str, maximum: int) -> Callable[[str], int]:
    def parse(value: str) -> int:
        try:
            parsed = int(value, 10)
        except ValueError as exc:
            raise argparse.ArgumentTypeError(f"{option} must be an integer") from exc
        if not 1 <= parsed <= maximum:
            raise argparse.ArgumentTypeError(
                f"{option} must be between 1 and {maximum} (0/unlimited is not allowed)"
            )
        return parsed

    return parse


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Stream read-only Extended MAPI messages into the local RAGSearch service"
    )
    parser.add_argument(
        "--executable",
        type=Path,
        default=DEFAULT_EXECUTABLE,
        help="Path to the built NativeMapiProbe.exe",
    )
    parser.add_argument(
        "--service-url",
        default=os.environ.get("RAGSEARCH_SERVICE_URL", DEFAULT_SERVICE_URL),
        help="Plain HTTP loopback service base URL",
    )
    parser.add_argument(
        "--token-path",
        type=Path,
        default=Settings.default().token_path,
        help="Path to the existing RAGSearch service token",
    )
    parser.add_argument(
        "--spool-dir",
        type=Path,
        default=Settings.default().spool_dir,
        help=(
            "Base spool directory shared with the service. A private, short-lived "
            "native import subdirectory is created underneath it."
        ),
    )
    parser.add_argument(
        "--cancel-file",
        type=Path,
        help="Optional sentinel file whose existence requests graceful cancellation",
    )
    parser.add_argument(
        "--max-stores",
        type=_bounded_integer("--max-stores", MAX_STORES),
        default=None,
    )
    parser.add_argument(
        "--max-folders",
        type=_bounded_integer("--max-folders", MAX_FOLDERS),
        default=None,
    )
    parser.add_argument(
        "--max-messages",
        type=_bounded_integer("--max-messages", MAX_MESSAGES),
        default=None,
    )
    parser.add_argument(
        "--full-scan",
        action="store_true",
        help=(
            "Scan all stores, folders and messages (passes zero/unlimited native limits); "
            "cannot be combined with --max-stores/--max-folders/--max-messages"
        ),
    )
    parser.add_argument(
        "--body-preview-chars",
        type=_bounded_integer("--body-preview-chars", MAX_BODY_CHARS),
        default=None,
        help=(
            "Maximum body characters read per message (default: 200000 for a bounded "
            "probe, 4000000 for --full-scan)"
        ),
    )
    parser.add_argument(
        "--store-contains",
        help="Optional NativeMapiProbe store display-name filter",
    )
    return parser


def validate_executable(path: Path) -> Path:
    try:
        resolved = Path(path).resolve(strict=True)
    except (OSError, RuntimeError) as exc:
        raise ImportFailure(f"NativeMapiProbe executable does not exist: {path}") from exc
    if not resolved.is_file():
        raise ImportFailure(f"NativeMapiProbe executable is not a file: {resolved}")
    if resolved.name.casefold() != "nativemapiprobe.exe":
        raise ImportFailure(
            f"Native executable must be named NativeMapiProbe.exe: {resolved}"
        )
    return resolved


def validate_service_url(value: str) -> str:
    if not isinstance(value, str) or not value or any(ord(char) < 32 for char in value):
        raise ImportFailure("Service URL must be a plain HTTP loopback URL")
    try:
        parsed = urlsplit(value)
        port = parsed.port
    except ValueError as exc:
        raise ImportFailure("Service URL has an invalid port") from exc

    hostname = parsed.hostname
    loopback = hostname is not None and hostname.casefold() == "localhost"
    if hostname and not loopback:
        try:
            loopback = ipaddress.ip_address(hostname).is_loopback
        except ValueError:
            loopback = False

    if (
        parsed.scheme.casefold() != "http"
        or not loopback
        or parsed.username is not None
        or parsed.password is not None
        or parsed.query
        or parsed.fragment
        or parsed.path not in {"", "/"}
    ):
        raise ImportFailure(
            "Service URL must be a plain HTTP loopback base URL"
        )

    if port == 0:
        raise ImportFailure("Service URL port must be between 1 and 65535")
    port = 80 if port is None else port
    host_for_url = f"[{hostname}]" if ":" in hostname else hostname
    return f"http://{host_for_url}:{port}"


def read_existing_token(path: Path) -> str:
    candidate = Path(path)
    try:
        resolved = candidate.resolve(strict=True)
    except (OSError, RuntimeError) as exc:
        raise ImportFailure(f"Service token file does not exist: {candidate}") from exc
    if not resolved.is_file():
        raise ImportFailure(f"Service token path is not a file: {resolved}")
    try:
        token = resolved.read_text(encoding="ascii").strip()
    except (OSError, UnicodeError) as exc:
        raise ImportFailure(f"Could not read the service token: {resolved}") from exc
    if len(token) < 32 or any(not 0x21 <= ord(char) <= 0x7E for char in token):
        raise ImportFailure(f"Service token is missing or invalid: {resolved}")
    return token


def validate_spool_directory(path: Path) -> Path:
    candidate = Path(path)
    try:
        resolved = candidate.resolve(strict=True)
    except (OSError, RuntimeError) as exc:
        raise ImportFailure(f"Spool directory does not exist: {candidate}") from exc
    if not resolved.is_dir():
        raise ImportFailure(f"Spool path is not a directory: {resolved}")
    return resolved


def _is_strict_descendant(candidate: Path, root: Path) -> bool:
    try:
        common = os.path.commonpath(
            (os.path.normcase(str(root)), os.path.normcase(str(candidate)))
        )
    except ValueError:
        return False
    return common == os.path.normcase(str(root)) and candidate != root


def _create_owned_spool_run(base_dir: Path) -> OwnedSpoolRun:
    base_dir = validate_spool_directory(base_dir)
    try:
        created = Path(tempfile.mkdtemp(prefix=OWNED_RUN_PREFIX, dir=base_dir))
        run_dir = created.resolve(strict=True)
    except (OSError, RuntimeError) as exc:
        raise ImportFailure(f"Could not create a private native spool directory: {exc}") from exc

    if run_dir.parent != base_dir or not _is_strict_descendant(run_dir, base_dir):
        try:
            shutil.rmtree(run_dir)
        except OSError:
            pass
        raise ImportFailure("Private native spool directory escaped the configured spool")

    marker_token = secrets.token_hex(32)
    marker = run_dir / OWNERSHIP_MARKER_NAME
    try:
        marker.write_text(marker_token, encoding="ascii")
    except OSError as exc:
        try:
            shutil.rmtree(run_dir)
        except OSError:
            pass
        raise ImportFailure(f"Could not mark the private native spool directory: {exc}") from exc
    return OwnedSpoolRun(base_dir, run_dir, marker_token)


def _cleanup_owned_spool_run(owned: OwnedSpoolRun) -> None:
    if not owned.run_dir.exists():
        return
    try:
        current = owned.run_dir.resolve(strict=True)
    except (OSError, RuntimeError) as exc:
        raise ImportFailure("Could not resolve the private native spool during cleanup") from exc
    if (
        current != owned.run_dir
        or current.parent != owned.base_dir
        or not _is_strict_descendant(current, owned.base_dir)
    ):
        raise ImportFailure("Refusing to clean a changed or unowned native spool path")

    marker = current / OWNERSHIP_MARKER_NAME
    try:
        marker_mode = marker.lstat().st_mode
        marker_value = marker.read_text(encoding="ascii")
    except (OSError, UnicodeError) as exc:
        raise ImportFailure(
            "Refusing to clean a native spool without its ownership marker"
        ) from exc
    if not stat.S_ISREG(marker_mode) or marker.is_symlink() or not secrets.compare_digest(
        marker_value, owned.marker_token
    ):
        raise ImportFailure("Refusing to clean a native spool with an invalid ownership marker")

    try:
        shutil.rmtree(current)
    except OSError as exc:
        raise ImportFailure(f"Could not clean the private native spool directory: {exc}") from exc


def _safe_native_attachment_path(raw_path: str, spool_dir: Path) -> Path:
    if not raw_path or len(raw_path) > MAX_TEMP_PATH_CHARS or "\0" in raw_path:
        raise ImportFailure("Native attachment temp_path is invalid")
    try:
        candidate = Path(raw_path).resolve(strict=True)
        mode = candidate.stat().st_mode
    except (OSError, RuntimeError) as exc:
        raise ImportFailure(
            f"Native attachment temp_path does not exist: {raw_path}"
        ) from exc
    if not _is_strict_descendant(candidate, spool_dir) or not stat.S_ISREG(mode):
        raise ImportFailure(
            "Native attachment temp_path is outside the private spool or is not a regular file"
        )
    return candidate


def _delete_owned_attachment_files(paths: list[Path], spool_dir: Path) -> None:
    for expected in paths:
        if not expected.exists():
            continue
        try:
            current = expected.resolve(strict=True)
            mode = current.stat().st_mode
        except (OSError, RuntimeError):
            continue
        if current != expected or not _is_strict_descendant(current, spool_dir):
            continue
        if not stat.S_ISREG(mode):
            continue
        try:
            current.unlink()
        except OSError:
            # The final ownership-checked run-directory cleanup retries this file.
            pass


def write_progress(event: Mapping[str, object], *, stream: IO[str] | None = None) -> None:
    output = sys.stdout if stream is None else stream
    print(
        PROGRESS_PREFIX
        + json.dumps(dict(event), ensure_ascii=False, separators=(",", ":")),
        file=output,
        flush=True,
    )


def _report_progress(
    callback: ProgressCallback | None,
    *,
    phase: str,
    current: int,
    total: int,
    message: str,
    bodies_truncated: int = 0,
) -> None:
    if callback is not None:
        callback(
            {
                "phase": phase,
                "current": current,
                "total": total,
                "message": message,
                "bodies_truncated": bodies_truncated,
            }
        )


def _check_cancel(cancel_file: Path | None) -> None:
    if cancel_file is None:
        return
    try:
        cancelled = os.path.lexists(cancel_file)
    except (OSError, ValueError) as exc:
        raise ImportFailure(f"Could not inspect the cancellation sentinel: {exc}") from exc
    if cancelled:
        raise ImportCancelled("Native MAPI import cancelled")


def _native_text(
    record: Mapping[str, object],
    name: str,
    *,
    required: bool,
    maximum_chars: int,
) -> str:
    value = record.get(name)
    if not isinstance(value, str):
        raise ImportFailure(f"Native JSONL field {name!r} must be a string")
    if required and not value.strip():
        raise ImportFailure(f"Native JSONL field {name!r} must not be empty")
    if len(value) > maximum_chars:
        raise ImportFailure(
            f"Native JSONL field {name!r} exceeds {maximum_chars} characters"
        )
    if "\0" in value:
        raise ImportFailure(f"Native JSONL field {name!r} contains a null character")
    return value


def _native_timestamp(record: Mapping[str, object], name: str) -> str | None:
    value = record.get(name)
    if value is None:
        return None
    if not isinstance(value, str) or not value or len(value) > MAX_TIMESTAMP_CHARS:
        raise ImportFailure(
            f"Native JSONL field {name!r} must be an ISO-8601 string or null"
        )
    try:
        parsed = dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as exc:
        raise ImportFailure(
            f"Native JSONL field {name!r} must be an ISO-8601 string or null"
        ) from exc
    if "T" not in value or parsed.tzinfo is None or parsed.utcoffset() is None:
        raise ImportFailure(
            f"Native JSONL field {name!r} must include a date, time, and UTC offset"
        )
    return value


def _native_bool(record: Mapping[str, object], name: str) -> bool:
    value = record.get(name)
    if type(value) is not bool:
        raise ImportFailure(f"Native JSONL field {name!r} must be a boolean")
    return value


def _native_attachment(
    raw: object,
    *,
    attachment_index: int,
    spool_dir: Path,
    maximum_attachment_bytes: int,
) -> tuple[dict[str, object], Path | None]:
    if not isinstance(raw, dict):
        raise ImportFailure(
            f"Native JSONL attachments[{attachment_index}] must be an object"
        )
    name = _native_text(
        raw,
        "name",
        required=False,
        maximum_chars=MAX_ATTACHMENT_NAME_CHARS,
    )
    content_type = _native_text(
        raw,
        "content_type",
        required=False,
        maximum_chars=MAX_CONTENT_TYPE_CHARS,
    )
    temp_path = _native_text(
        raw,
        "temp_path",
        required=False,
        maximum_chars=MAX_TEMP_PATH_CHARS,
    )
    size = raw.get("size")
    if (
        type(size) is not int
        or not 0 <= size <= MAX_DECLARED_ATTACHMENT_BYTES
    ):
        raise ImportFailure(
            f"Native JSONL attachments[{attachment_index}].size must be a bounded "
            "non-negative integer"
        )

    safe_path: Path | None = None
    if temp_path:
        safe_path = _safe_native_attachment_path(temp_path, spool_dir)
        try:
            actual_size = safe_path.stat().st_size
        except OSError as exc:
            raise ImportFailure(
                f"Could not stat Native JSONL attachments[{attachment_index}].temp_path"
            ) from exc
        if actual_size > maximum_attachment_bytes:
            raise ImportFailure(
                f"Native JSONL attachments[{attachment_index}] exceeds the spool file limit"
            )
        temp_path = str(safe_path)

    return (
        {
            "name": name,
            "size": size,
            "content_type": content_type,
            "temp_path": temp_path,
        },
        safe_path,
    )


def map_native_message(
    record: object,
    *,
    maximum_body_chars: int,
    spool_dir: Path,
    maximum_attachment_bytes: int = Settings.default().attachment_limit_bytes,
) -> dict[str, object]:
    if not isinstance(record, dict):
        raise ImportFailure("Each NativeMapiProbe JSONL record must be an object")
    if type(maximum_body_chars) is not int or not 1 <= maximum_body_chars <= MAX_BODY_CHARS:
        raise ImportFailure(f"maximum_body_chars must be between 1 and {MAX_BODY_CHARS}")
    if type(maximum_attachment_bytes) is not int or maximum_attachment_bytes < 1:
        raise ImportFailure("maximum_attachment_bytes must be a positive integer")
    spool_dir = validate_spool_directory(spool_dir)

    store_id = _native_text(
        record, "store_id", required=True, maximum_chars=MAX_ID_CHARS
    )
    entry_id = _native_text(
        record, "entry_id", required=True, maximum_chars=MAX_ID_CHARS
    )
    folder_entry_id = _native_text(
        record, "folder_entry_id", required=True, maximum_chars=MAX_ID_CHARS
    )
    folder_path = _native_text(
        record, "folder_path", required=True, maximum_chars=MAX_FOLDER_PATH_CHARS
    )
    store_name = _native_text(
        record, "store_name", required=False, maximum_chars=MAX_STORE_NAME_CHARS
    )
    subject = _native_text(
        record, "subject", required=False, maximum_chars=MAX_SUBJECT_CHARS
    )
    body = _native_text(
        record, "body", required=False, maximum_chars=maximum_body_chars
    )
    if len(body) > maximum_body_chars:
        raise ImportFailure("NativeMapiProbe emitted a body larger than the requested limit")
    body_available = _native_bool(record, "body_available")
    body_truncated = _native_bool(record, "body_truncated")
    if not body_available and body:
        raise ImportFailure("Native JSONL body must be empty when body_available is false")
    if body_truncated and not body_available:
        raise ImportFailure("Native JSONL body cannot be truncated when it is unavailable")

    raw_attachments = record.get("attachments")
    if not isinstance(raw_attachments, list):
        raise ImportFailure("Native JSONL field 'attachments' must be an array")
    if len(raw_attachments) > MAX_ATTACHMENTS_PER_MESSAGE:
        raise ImportFailure(
            f"Native JSONL field 'attachments' exceeds {MAX_ATTACHMENTS_PER_MESSAGE} items"
        )
    attachments: list[dict[str, object]] = []
    for index, raw_attachment in enumerate(raw_attachments):
        attachment, _ = _native_attachment(
            raw_attachment,
            attachment_index=index,
            spool_dir=spool_dir,
            maximum_attachment_bytes=maximum_attachment_bytes,
        )
        attachments.append(attachment)

    return {
        "entry_id": entry_id,
        "store_id": store_id,
        "folder_entry_id": folder_entry_id,
        "folder_path": folder_path,
        "store_name": store_name,
        "subject": subject,
        "sender_name": _native_text(
            record,
            "sender_name",
            required=False,
            maximum_chars=MAX_METADATA_FIELD_CHARS,
        ),
        "sender_email": _native_text(
            record,
            "sender_email",
            required=False,
            maximum_chars=MAX_METADATA_FIELD_CHARS,
        ),
        "to": _native_text(
            record,
            "to",
            required=False,
            maximum_chars=MAX_ADDRESS_FIELD_CHARS,
        ),
        "cc": _native_text(
            record,
            "cc",
            required=False,
            maximum_chars=MAX_ADDRESS_FIELD_CHARS,
        ),
        "sent_at": _native_timestamp(record, "sent_at"),
        "received_at": _native_timestamp(record, "received_at"),
        "modified_at": _native_timestamp(record, "modified_at"),
        "internet_message_id": _native_text(
            record,
            "internet_message_id",
            required=False,
            maximum_chars=MAX_METADATA_FIELD_CHARS,
        ),
        "conversation_id": _native_text(
            record,
            "conversation_id",
            required=False,
            maximum_chars=MAX_METADATA_FIELD_CHARS,
        ),
        "body": body,
        "attachments": attachments,
    }


def _read_response_body(response: object) -> bytes:
    body = response.read(MAX_HTTP_RESPONSE_BYTES + 1)  # type: ignore[attr-defined]
    if len(body) > MAX_HTTP_RESPONSE_BYTES:
        raise ImportFailure("RAGSearch service response exceeded the safety limit")
    return body


def post_message(
    endpoint: str,
    token: str,
    message: Mapping[str, object],
    *,
    opener: Callable[..., object] = urllib.request.urlopen,
) -> None:
    encoded = json.dumps(
        {"messages": [message]},
        ensure_ascii=False,
        separators=(",", ":"),
    ).encode("utf-8")
    if len(encoded) > MAX_HTTP_REQUEST_BYTES:
        raise ImportFailure("Native message exceeds the local HTTP request safety limit")
    request = urllib.request.Request(
        endpoint,
        data=encoded,
        method="POST",
        headers={
            "Content-Type": "application/json; charset=utf-8",
            "Accept": "application/json",
            TOKEN_HEADER: token,
        },
    )
    try:
        with opener(  # type: ignore[attr-defined]
            request, timeout=HTTP_TIMEOUT_SECONDS
        ) as response:
            status = getattr(response, "status", None)
            if status is None:
                status = response.getcode()  # type: ignore[attr-defined]
            raw_response = _read_response_body(response)
    except urllib.error.HTTPError as exc:
        detail = exc.read(MAX_HTTP_RESPONSE_BYTES).decode("utf-8", errors="replace").strip()
        raise ImportFailure(
            f"RAGSearch service returned HTTP {exc.code}: {detail[:512]}"
        ) from exc
    except urllib.error.URLError as exc:
        raise ImportFailure(f"Could not reach the local RAGSearch service: {exc.reason}") from exc
    except (TimeoutError, OSError) as exc:
        raise ImportFailure(f"Could not reach the local RAGSearch service: {exc}") from exc

    if status != 200:
        raise ImportFailure(f"RAGSearch service returned unexpected HTTP status {status}")
    try:
        outcome = json.loads(raw_response.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ImportFailure("RAGSearch service returned invalid UTF-8 JSON") from exc
    if (
        not isinstance(outcome, dict)
        or type(outcome.get("accepted")) is not int
        or type(outcome.get("failed")) is not int
        or outcome["accepted"] != 1
        or outcome["failed"] != 0
    ):
        raise ImportFailure("RAGSearch service rejected the native message")


def _stop_process(process: subprocess.Popen[str]) -> None:
    if process.poll() is not None:
        return
    process.terminate()
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait()


def import_native_messages(
    *,
    executable: Path,
    service_url: str,
    token_path: Path,
    spool_dir: Path,
    max_stores: int | None,
    max_folders: int | None,
    max_messages: int | None,
    body_preview_chars: int,
    full_scan: bool = False,
    store_contains: str | None = None,
    cancel_file: Path | None = None,
    maximum_attachment_bytes: int = Settings.default().attachment_limit_bytes,
    progress: ProgressCallback | None = None,
    popen: Callable[..., subprocess.Popen[str]] = subprocess.Popen,
    opener: Callable[..., object] = urllib.request.urlopen,
) -> ImportResult:
    if type(full_scan) is not bool:
        raise ImportFailure("full_scan must be a boolean")
    numeric_limits = (
        ("max_stores", max_stores, MAX_STORES),
        ("max_folders", max_folders, MAX_FOLDERS),
        ("max_messages", max_messages, MAX_MESSAGES),
    )
    if full_scan:
        if any(value is not None for _, value, _ in numeric_limits):
            raise ImportFailure(
                "full_scan cannot be combined with max_stores, max_folders, or max_messages"
            )
        probe_max_stores = probe_max_folders = probe_max_messages = 0
    else:
        for name, value, maximum in numeric_limits:
            if type(value) is not int or not 1 <= value <= maximum:
                raise ImportFailure(f"{name} must be between 1 and {maximum}")
        probe_max_stores = max_stores
        probe_max_folders = max_folders
        probe_max_messages = max_messages
    if type(body_preview_chars) is not int or not 1 <= body_preview_chars <= MAX_BODY_CHARS:
        raise ImportFailure(f"body_preview_chars must be between 1 and {MAX_BODY_CHARS}")
    if type(maximum_attachment_bytes) is not int or maximum_attachment_bytes < 1:
        raise ImportFailure("maximum_attachment_bytes must be a positive integer")

    executable = validate_executable(executable)
    base_url = validate_service_url(service_url)
    token = read_existing_token(token_path)
    spool_base = validate_spool_directory(spool_dir)

    if cancel_file is not None:
        if "\0" in str(cancel_file):
            raise ImportFailure("--cancel-file contains a null character")
        try:
            cancel_file = Path(cancel_file).absolute()
        except (OSError, RuntimeError) as exc:
            raise ImportFailure(f"Could not resolve --cancel-file: {exc}") from exc

    if store_contains is not None:
        if (
            not store_contains.strip()
            or len(store_contains) > MAX_STORE_FILTER_CHARS
            or "\0" in store_contains
        ):
            raise ImportFailure(
                f"--store-contains must contain 1 to {MAX_STORE_FILTER_CHARS} characters"
            )

    owned = _create_owned_spool_run(spool_base)
    process: subprocess.Popen[str] | None = None
    stdout: IO[str] | None = None
    imported = 0
    bodies_truncated = 0
    progress_total = 0 if full_scan else int(probe_max_messages)
    endpoint = base_url + "/v1/messages"
    # C++ JSON escaping can expand one body character to six ASCII characters;
    # bounded metadata/attachment descriptors fit in the fixed overhead allowance.
    maximum_line_chars = body_preview_chars * 6 + MAX_JSONL_OVERHEAD_CHARS
    active_error: BaseException | None = None
    try:
        _check_cancel(cancel_file)
        command = [
            str(executable),
            "--jsonl",
            "--max-stores",
            str(probe_max_stores),
            "--max-folders",
            str(probe_max_folders),
            "--max-messages",
            str(probe_max_messages),
            "--body-preview-chars",
            str(body_preview_chars),
            "--spool-dir",
            str(owned.run_dir),
            "--max-attachment-bytes",
            str(maximum_attachment_bytes),
            "--max-total-attachment-bytes",
            "0",
        ]
        if store_contains is not None:
            command.extend(("--store-contains", store_contains))

        _report_progress(
            progress,
            phase="starting",
            current=0,
            total=progress_total,
            message="Starting native MAPI scan",
        )
        try:
            process = popen(
                command,
                stdout=subprocess.PIPE,
                stderr=None,  # Native diagnostics remain visible without buffering in Python.
                text=True,
                encoding="utf-8",
                errors="strict",
                bufsize=1,
                shell=False,
            )
        except OSError as exc:
            raise ImportFailure(f"Could not start NativeMapiProbe.exe: {exc}") from exc
        if process.stdout is None:
            _stop_process(process)
            raise ImportFailure("Could not capture NativeMapiProbe JSONL output")
        stdout = process.stdout

        output_line = 0
        while True:
            _check_cancel(cancel_file)
            try:
                line = stdout.readline(maximum_line_chars + 1)
            except UnicodeError as exc:
                raise ImportFailure("NativeMapiProbe stdout is not valid UTF-8") from exc
            if not line:
                break
            output_line += 1
            if len(line) > maximum_line_chars:
                raise ImportFailure("NativeMapiProbe emitted an oversized JSONL record")
            if not line.strip():
                continue
            if probe_max_messages and imported >= probe_max_messages:
                raise ImportFailure("NativeMapiProbe emitted more messages than requested")
            if imported == 0 and line.startswith("\ufeff"):
                line = line.removeprefix("\ufeff")
            try:
                native_record = json.loads(line)
            except json.JSONDecodeError as exc:
                raise ImportFailure(
                    f"NativeMapiProbe emitted invalid JSONL at output line {output_line}"
                ) from exc
            message = map_native_message(
                native_record,
                maximum_body_chars=body_preview_chars,
                spool_dir=owned.run_dir,
                maximum_attachment_bytes=maximum_attachment_bytes,
            )
            owned_files = [
                Path(str(attachment["temp_path"]))
                for attachment in message["attachments"]  # type: ignore[index]
                if attachment["temp_path"]  # type: ignore[index]
            ]
            try:
                _check_cancel(cancel_file)
                post_message(endpoint, token, message, opener=opener)
                imported += 1
                if native_record["body_truncated"]:
                    bodies_truncated += 1
                _report_progress(
                    progress,
                    phase="importing",
                    current=imported,
                    total=progress_total,
                    message=(
                        f"Imported {imported} message(s); "
                        f"truncated bodies: {bodies_truncated}"
                    ),
                    bodies_truncated=bodies_truncated,
                )
                _check_cancel(cancel_file)
            finally:
                _delete_owned_attachment_files(owned_files, owned.run_dir)

        probe_exit_code = process.wait()
        _report_progress(
            progress,
            phase="complete" if probe_exit_code == 0 else "failed",
            current=imported,
            total=imported if probe_exit_code == 0 else progress_total,
            message=(
                f"Imported {imported} message(s); truncated bodies: {bodies_truncated}"
                if probe_exit_code == 0
                else f"NativeMapiProbe exited with code {probe_exit_code}"
            ),
            bodies_truncated=bodies_truncated,
        )
        return ImportResult(
            imported=imported,
            probe_exit_code=probe_exit_code,
            bodies_truncated=bodies_truncated,
        )
    except (ImportCancelled, KeyboardInterrupt) as exc:
        active_error = exc
        if process is not None:
            _stop_process(process)
        _report_progress(
            progress,
            phase="cancelled",
            current=imported,
            total=progress_total,
            message="Native MAPI import cancelled",
            bodies_truncated=bodies_truncated,
        )
        raise
    except BaseException as exc:
        active_error = exc
        if process is not None:
            _stop_process(process)
        _report_progress(
            progress,
            phase="failed",
            current=imported,
            total=progress_total,
            message="Native MAPI import failed",
            bodies_truncated=bodies_truncated,
        )
        raise
    finally:
        if stdout is not None:
            stdout.close()
        try:
            _cleanup_owned_spool_run(owned)
        except ImportFailure as cleanup_error:
            if active_error is not None:
                cleanup_error.add_note(f"Original import failure: {active_error}")
                raise cleanup_error from active_error
            raise


def main(argv: list[str] | None = None) -> int:
    parser = _parser()
    args = parser.parse_args(argv)
    explicit_numeric_limits = any(
        value is not None
        for value in (args.max_stores, args.max_folders, args.max_messages)
    )
    if args.full_scan and explicit_numeric_limits:
        parser.error(
            "--full-scan cannot be combined with --max-stores, --max-folders, "
            "or --max-messages"
        )
    if not args.full_scan:
        args.max_stores = (
            DEFAULT_MAX_STORES if args.max_stores is None else args.max_stores
        )
        args.max_folders = (
            DEFAULT_MAX_FOLDERS if args.max_folders is None else args.max_folders
        )
        args.max_messages = (
            DEFAULT_MAX_MESSAGES if args.max_messages is None else args.max_messages
        )
    args.body_preview_chars = (
        (MAX_BODY_CHARS if args.full_scan else DEFAULT_BODY_CHARS)
        if args.body_preview_chars is None
        else args.body_preview_chars
    )
    try:
        result = import_native_messages(
            executable=args.executable,
            service_url=args.service_url,
            token_path=args.token_path,
            spool_dir=args.spool_dir,
            max_stores=args.max_stores,
            max_folders=args.max_folders,
            max_messages=args.max_messages,
            body_preview_chars=args.body_preview_chars,
            full_scan=args.full_scan,
            store_contains=args.store_contains,
            cancel_file=args.cancel_file,
            progress=write_progress,
        )
    except ImportCancelled:
        print("Native MAPI import cancelled", file=sys.stderr)
        return 130
    except ImportFailure as exc:
        print(f"Native MAPI import failed: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("Native MAPI import interrupted", file=sys.stderr)
        return 130

    if result.probe_exit_code != 0:
        print(
            "NativeMapiProbe exited with code "
            f"{result.probe_exit_code} after importing {result.imported} message(s); "
            f"truncated bodies: {result.bodies_truncated}",
            file=sys.stderr,
        )
        return result.probe_exit_code
    print(
        f"Imported {result.imported} message(s) from NativeMapiProbe; "
        f"truncated bodies: {result.bodies_truncated}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
