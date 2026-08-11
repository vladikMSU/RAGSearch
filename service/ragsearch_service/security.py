from __future__ import annotations

import os
import secrets
import stat
import subprocess
import csv
from pathlib import Path


def _current_windows_sid() -> str:
    completed = subprocess.run(
        ["whoami", "/user", "/fo", "csv", "/nh"],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=10,
    )
    row = next(csv.reader(completed.stdout.splitlines()))
    if len(row) < 2 or not row[1].upper().startswith("S-1-"):
        raise RuntimeError("Could not determine the current Windows user SID")
    return row[1]


def ensure_private_path(path: Path) -> None:
    """Remove inherited Windows ACEs or apply a private POSIX mode."""
    path = Path(path)
    if os.name != "nt":
        mode = stat.S_IRWXU if path.is_dir() else stat.S_IRUSR | stat.S_IWUSR
        os.chmod(path, mode)
        return

    sid = _current_windows_sid()
    inheritance = "(OI)(CI)F" if path.is_dir() else "F"
    completed = subprocess.run(
        [
            "icacls",
            str(path),
            "/inheritance:r",
            "/grant:r",
            f"*{sid}:{inheritance}",
            f"*S-1-5-18:{inheritance}",
            f"*S-1-5-32-544:{inheritance}",
        ],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=15,
    )
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout).strip()
        raise RuntimeError(f"Could not secure local RAGSearch ACL for {path}: {detail}")


def load_or_create_token(path: Path) -> str:
    """Return a stable local bearer token, creating it atomically when absent."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    try:
        token = path.read_text(encoding="ascii").strip()
    except FileNotFoundError:
        token = secrets.token_urlsafe(32)
        flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL
        try:
            descriptor = os.open(path, flags, 0o600)
        except FileExistsError:
            token = path.read_text(encoding="ascii").strip()
        else:
            with os.fdopen(descriptor, "w", encoding="ascii", newline="\n") as stream:
                stream.write(token + "\n")
            ensure_private_path(path)

    if len(token) < 32:
        raise RuntimeError(f"Service token is missing or invalid: {path}")
    ensure_private_path(path)
    return token


def token_matches(expected: str, supplied: str | None) -> bool:
    if not supplied:
        return False
    return secrets.compare_digest(expected, supplied.strip())
