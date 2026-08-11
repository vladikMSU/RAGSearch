from __future__ import annotations

import io
import json
import subprocess
import tempfile
import threading
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from typing import Mapping
from unittest import mock

import import_native_mapi as adapter
from ragsearch_service.app import SearchService
from ragsearch_service.config import Settings
from ragsearch_service.http_api import create_http_server


class FakeResponse:
    def __init__(self, payload: object, *, status: int = 200) -> None:
        self.status = status
        self._body = json.dumps(payload).encode("utf-8")

    def __enter__(self) -> "FakeResponse":
        return self

    def __exit__(self, *args: object) -> None:
        return None

    def getcode(self) -> int:
        return self.status

    def read(self, amount: int = -1) -> bytes:
        return self._body[:amount] if amount >= 0 else self._body


class FakeProcess:
    def __init__(self, output: str, *, returncode: int = 0) -> None:
        self.stdout = io.StringIO(output)
        self.returncode = returncode
        self.terminated = False
        self.killed = False

    def poll(self) -> int | None:
        return self.returncode if self.terminated else None

    def terminate(self) -> None:
        self.terminated = True

    def kill(self) -> None:
        self.killed = True
        self.terminated = True

    def wait(self, timeout: float | None = None) -> int:
        return self.returncode


class NativeMapiAdapterTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        root = Path(self.temporary.name)
        self.executable = root / "NativeMapiProbe.exe"
        self.executable.write_bytes(b"test executable placeholder")
        self.token_path = root / "service-token"
        self.token = "t" * 43
        self.token_path.write_text(self.token + "\n", encoding="ascii")
        self.spool_dir = root / "spool"
        self.spool_dir.mkdir()

    def native_message(self, **changes: object) -> dict[str, object]:
        message: dict[str, object] = {
            "store_id": "store-1",
            "entry_id": "entry-1",
            "folder_entry_id": "folder-1",
            "folder_path": "Mailbox/Inbox",
            "store_name": "Mailbox",
            "subject": "Тема",
            "body": "Тело сообщения",
            "body_available": True,
            "body_truncated": False,
            "sender_name": "Алексей",
            "sender_email": "alex@example.test",
            "to": "user@example.test",
            "cc": "team@example.test",
            "sent_at": "2026-08-11T08:59:00Z",
            "received_at": "2026-08-11T09:00:00+00:00",
            "modified_at": None,
            "internet_message_id": "<entry-1@example.test>",
            "conversation_id": "conversation-1",
            "attachments": [],
        }
        message.update(changes)
        return message

    def import_arguments(self, **changes: object) -> dict[str, object]:
        arguments: dict[str, object] = {
            "executable": self.executable,
            "service_url": "http://127.0.0.1:8765",
            "token_path": self.token_path,
            "spool_dir": self.spool_dir,
            "max_stores": 1,
            "max_folders": 10,
            "max_messages": 1,
            "body_preview_chars": 1_000,
        }
        arguments.update(changes)
        return arguments

    def test_maps_complete_native_contract_and_attachment(self) -> None:
        attachment_path = self.spool_dir / "notes.txt"
        attachment_path.write_text("Содержимое вложения", encoding="utf-8")
        record = self.native_message(
            attachments=[
                {
                    "name": "notes.txt",
                    "size": attachment_path.stat().st_size,
                    "content_type": "text/plain",
                    "temp_path": str(attachment_path),
                },
                {
                    "name": "large.bin",
                    "size": 999_999_999,
                    "content_type": "application/octet-stream",
                    "temp_path": "",
                },
            ]
        )

        mapped = adapter.map_native_message(
            record,
            maximum_body_chars=1_000,
            spool_dir=self.spool_dir,
        )

        for field in (
            "store_id",
            "entry_id",
            "folder_entry_id",
            "folder_path",
            "store_name",
            "subject",
            "body",
            "sender_name",
            "sender_email",
            "to",
            "cc",
            "sent_at",
            "received_at",
            "modified_at",
            "internet_message_id",
            "conversation_id",
        ):
            self.assertEqual(record[field], mapped[field], field)
        attachments = mapped["attachments"]
        self.assertEqual(2, len(attachments))
        self.assertEqual(str(attachment_path.resolve()), attachments[0]["temp_path"])
        self.assertEqual("", attachments[1]["temp_path"])

    def test_native_contract_rejects_missing_fields_bad_dates_and_bounds(self) -> None:
        invalid_records = [
            (
                {
                    key: value
                    for key, value in self.native_message().items()
                    if key != "folder_entry_id"
                },
                "folder_entry_id",
            ),
            (self.native_message(entry_id=""), "entry_id"),
            (self.native_message(sent_at="2026-08-11 09:00:00"), "UTC offset"),
            (self.native_message(received_at=123), "ISO-8601"),
            (self.native_message(attachments={}), "array"),
            (
                self.native_message(
                    attachments=[
                        {
                            "name": "bad.txt",
                            "size": -1,
                            "content_type": "text/plain",
                            "temp_path": "",
                        }
                    ]
                ),
                "bounded",
            ),
        ]
        for record, error in invalid_records:
            with self.subTest(error=error), self.assertRaisesRegex(
                adapter.ImportFailure, error
            ):
                adapter.map_native_message(
                    record,
                    maximum_body_chars=1_000,
                    spool_dir=self.spool_dir,
                )

        with self.assertRaisesRegex(adapter.ImportFailure, "exceeds"):
            adapter.map_native_message(
                self.native_message(body="x" * 11),
                maximum_body_chars=10,
                spool_dir=self.spool_dir,
            )

    def test_attachment_path_must_be_existing_regular_file_inside_spool(self) -> None:
        outside = Path(self.temporary.name) / "outside.txt"
        outside.write_text("not owned", encoding="utf-8")
        directory = self.spool_dir / "directory"
        directory.mkdir()
        for raw_path in (outside, directory, self.spool_dir / "missing.txt"):
            record = self.native_message(
                attachments=[
                    {
                        "name": "unsafe.txt",
                        "size": 1,
                        "content_type": "text/plain",
                        "temp_path": str(raw_path),
                    }
                ]
            )
            with self.subTest(path=raw_path), self.assertRaises(adapter.ImportFailure):
                adapter.map_native_message(
                    record,
                    maximum_body_chars=1_000,
                    spool_dir=self.spool_dir,
                )
        self.assertTrue(outside.exists())

    def test_loopback_url_validation_rejects_nonlocal_or_credentialed_urls(self) -> None:
        self.assertEqual(
            "http://127.0.0.1:8765",
            adapter.validate_service_url("http://127.0.0.1:8765/"),
        )
        self.assertEqual(
            "http://[::1]:8765",
            adapter.validate_service_url("http://[::1]:8765"),
        )
        self.assertEqual(
            "http://localhost:80",
            adapter.validate_service_url("http://localhost"),
        )
        for unsafe in (
            "https://127.0.0.1:8765",
            "http://example.test:8765",
            "http://user:secret@127.0.0.1:8765",
            "http://127.0.0.1:8765/prefix",
            "http://127.0.0.1:8765?token=secret",
            "http://127.0.0.1:0",
        ):
            with self.subTest(url=unsafe), self.assertRaises(adapter.ImportFailure):
                adapter.validate_service_url(unsafe)

    def test_streams_one_post_per_record_with_private_spool_and_progress(self) -> None:
        records = [
            self.native_message(body_truncated=True),
            self.native_message(entry_id="entry-2", subject="Second"),
        ]
        process = FakeProcess(
            "".join(json.dumps(record, ensure_ascii=False) + "\n" for record in records),
            returncode=7,
        )
        popen_calls: list[tuple[list[str], dict[str, object]]] = []
        requests: list[object] = []
        progress: list[Mapping[str, object]] = []
        foreign = self.spool_dir / "foreign.txt"
        foreign.write_text("must survive", encoding="utf-8")

        def fake_popen(command: list[str], **kwargs: object) -> FakeProcess:
            popen_calls.append((command, kwargs))
            native_spool = Path(command[command.index("--spool-dir") + 1])
            self.assertTrue(native_spool.is_dir())
            self.assertEqual(self.spool_dir.resolve(), native_spool.parent)
            return process

        def fake_open(request: object, *, timeout: int) -> FakeResponse:
            requests.append(request)
            self.assertEqual(adapter.HTTP_TIMEOUT_SECONDS, timeout)
            return FakeResponse({"accepted": 1, "failed": 0, "errors": []})

        result = adapter.import_native_messages(
            **self.import_arguments(
                service_url="http://localhost:8765",
                max_stores=2,
                max_folders=20,
                max_messages=2,
            ),
            store_contains="Archive",
            progress=progress.append,
            popen=fake_popen,
            opener=fake_open,
        )

        self.assertEqual(
            adapter.ImportResult(imported=2, probe_exit_code=7, bodies_truncated=1),
            result,
        )
        command, options = popen_calls[0]
        self.assertEqual(str(self.executable.resolve()), command[0])
        for option in (
            "--jsonl",
            "--max-stores",
            "--max-folders",
            "--max-messages",
            "--body-preview-chars",
            "--spool-dir",
            "--max-attachment-bytes",
            "--max-total-attachment-bytes",
        ):
            self.assertIn(option, command)
        native_spool = Path(command[command.index("--spool-dir") + 1])
        self.assertFalse(native_spool.exists())
        self.assertTrue(foreign.exists())
        self.assertEqual(False, options["shell"])
        self.assertIsNone(options["stderr"])
        self.assertEqual(subprocess.PIPE, options["stdout"])

        self.assertEqual(2, len(requests))
        first = requests[0]
        self.assertEqual("http://localhost:8765/v1/messages", first.full_url)
        self.assertEqual("POST", first.get_method())
        request_headers = {name.casefold(): value for name, value in first.header_items()}
        self.assertEqual(self.token, request_headers[adapter.TOKEN_HEADER.casefold()])
        first_payload = json.loads(first.data.decode("utf-8"))
        self.assertEqual("folder-1", first_payload["messages"][0]["folder_entry_id"])
        self.assertEqual(
            ["starting", "importing", "importing", "failed"],
            [event["phase"] for event in progress],
        )
        self.assertEqual(1, progress[-1]["bodies_truncated"])

    def test_service_rejection_stops_probe_and_cleans_only_owned_run(self) -> None:
        process = FakeProcess("")
        native_file: Path | None = None
        foreign = self.spool_dir / "foreign.txt"
        foreign.write_text("must survive", encoding="utf-8")

        def fake_popen(command: list[str], **kwargs: object) -> FakeProcess:
            nonlocal native_file
            run_dir = Path(command[command.index("--spool-dir") + 1])
            native_file = run_dir / "message" / "notes.txt"
            native_file.parent.mkdir()
            native_file.write_text("sensitive", encoding="utf-8")
            record = self.native_message(
                attachments=[
                    {
                        "name": "notes.txt",
                        "size": native_file.stat().st_size,
                        "content_type": "text/plain",
                        "temp_path": str(native_file),
                    }
                ]
            )
            process.stdout = io.StringIO(json.dumps(record) + "\n")
            return process

        with self.assertRaisesRegex(adapter.ImportFailure, "rejected"):
            adapter.import_native_messages(
                **self.import_arguments(),
                popen=fake_popen,
                opener=lambda *args, **kwargs: FakeResponse(
                    {"accepted": 0, "failed": 1, "errors": [{"error": "bad"}]}
                ),
            )

        self.assertTrue(process.terminated)
        self.assertIsNotNone(native_file)
        self.assertFalse(native_file.parent.parent.exists())
        self.assertTrue(foreign.exists())

    def test_streamed_full_message_and_attachment_match_live_service(self) -> None:
        settings = Settings.explicit(
            Path(self.temporary.name) / "live-service",
            spool_dir=self.spool_dir,
            port=0,
        )
        service = SearchService(settings)
        server = create_http_server(service)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(thread.join, 5)
        self.addCleanup(server.server_close)
        self.addCleanup(server.shutdown)
        process = FakeProcess("")
        native_spool: Path | None = None

        def fake_popen(command: list[str], **kwargs: object) -> FakeProcess:
            nonlocal native_spool
            native_spool = Path(command[command.index("--spool-dir") + 1])
            attachment = native_spool / "message-1" / "notes.txt"
            attachment.parent.mkdir()
            attachment.write_text("УникальныйМаркерВложения", encoding="utf-8")
            record = self.native_message(
                body="УникальныйНативныйМаркер",
                attachments=[
                    {
                        "name": "notes.txt",
                        "size": attachment.stat().st_size,
                        "content_type": "text/plain",
                        "temp_path": str(attachment),
                    }
                ],
            )
            process.stdout = io.StringIO(json.dumps(record, ensure_ascii=False) + "\n")
            return process

        result = adapter.import_native_messages(
            **self.import_arguments(
                service_url=f"http://127.0.0.1:{server.server_port}",
                token_path=settings.token_path,
            ),
            popen=fake_popen,
        )

        self.assertEqual(adapter.ImportResult(imported=1, probe_exit_code=0), result)
        self.assertIsNotNone(native_spool)
        self.assertFalse(native_spool.exists())
        found = service.search({"query": "УникальныйМаркерВложения", "limit": 1})[
            "results"
        ][0]
        self.assertEqual("store-1", found["store_id"])
        self.assertEqual("entry-1", found["entry_id"])
        self.assertEqual("folder-1", found["folder_entry_id"])
        self.assertIn("attachment:notes.txt", found["matched_sources"])
        with service.database.session() as connection:
            stored = connection.execute(
                "SELECT store_name, sender_email, received_at FROM messages "
                "WHERE store_id = ? AND entry_id = ?",
                ("store-1", "entry-1"),
            ).fetchone()
        self.assertEqual(
            ("Mailbox", "alex@example.test", "2026-08-11T09:00:00+00:00"),
            tuple(stored),
        )

    def test_full_scan_passes_zero_limits_and_cannot_mix_bounded_limits(self) -> None:
        commands: list[list[str]] = []

        def fake_popen(command: list[str], **kwargs: object) -> FakeProcess:
            commands.append(command)
            return FakeProcess("")

        result = adapter.import_native_messages(
            **self.import_arguments(
                max_stores=None,
                max_folders=None,
                max_messages=None,
                full_scan=True,
            ),
            popen=fake_popen,
        )
        self.assertEqual(adapter.ImportResult(0, 0), result)
        command = commands[0]
        for option in ("--max-stores", "--max-folders", "--max-messages"):
            self.assertEqual("0", command[command.index(option) + 1])

        with self.assertRaisesRegex(adapter.ImportFailure, "cannot be combined"):
            adapter.import_native_messages(
                **self.import_arguments(full_scan=True),
                popen=fake_popen,
            )

    def test_cancel_sentinel_avoids_starting_probe_and_cleans_run(self) -> None:
        cancel_file = Path(self.temporary.name) / "cancel"
        cancel_file.write_text("cancel", encoding="ascii")
        progress: list[Mapping[str, object]] = []
        with self.assertRaises(adapter.ImportCancelled):
            adapter.import_native_messages(
                **self.import_arguments(),
                cancel_file=cancel_file,
                progress=progress.append,
                popen=mock.Mock(side_effect=AssertionError("must not start")),
            )
        self.assertEqual(["cancelled"], [event["phase"] for event in progress])
        self.assertFalse(
            any(
                child.name.startswith(adapter.OWNED_RUN_PREFIX)
                for child in self.spool_dir.iterdir()
            )
        )

    def test_progress_protocol_is_one_prefixed_json_object_per_line(self) -> None:
        stream = io.StringIO()
        adapter.write_progress(
            {"phase": "importing", "current": 7, "total": 0, "message": "ok"},
            stream=stream,
        )
        line = stream.getvalue()
        self.assertTrue(line.startswith(adapter.PROGRESS_PREFIX))
        self.assertTrue(line.endswith("\n"))
        self.assertEqual(
            {"phase": "importing", "current": 7, "total": 0, "message": "ok"},
            json.loads(line.removeprefix(adapter.PROGRESS_PREFIX)),
        )

    def test_main_propagates_native_exit_code_and_progress_callback(self) -> None:
        with redirect_stderr(io.StringIO()), redirect_stdout(io.StringIO()), mock.patch.object(
            adapter,
            "import_native_messages",
            return_value=adapter.ImportResult(imported=3, probe_exit_code=17),
        ) as importer:
            exit_code = adapter.main([])

        self.assertEqual(17, exit_code)
        self.assertIs(importer.call_args.kwargs["progress"], adapter.write_progress)
        self.assertEqual(adapter.DEFAULT_MAX_MESSAGES, importer.call_args.kwargs["max_messages"])
        self.assertEqual(
            adapter.DEFAULT_BODY_CHARS,
            importer.call_args.kwargs["body_preview_chars"],
        )

        with redirect_stderr(io.StringIO()), redirect_stdout(io.StringIO()), mock.patch.object(
            adapter,
            "import_native_messages",
            return_value=adapter.ImportResult(imported=0, probe_exit_code=0),
        ) as full_importer:
            self.assertEqual(0, adapter.main(["--full-scan"]))
        self.assertTrue(full_importer.call_args.kwargs["full_scan"])
        self.assertIsNone(full_importer.call_args.kwargs["max_messages"])
        self.assertEqual(
            adapter.MAX_BODY_CHARS,
            full_importer.call_args.kwargs["body_preview_chars"],
        )

    def test_cli_rejects_unlimited_numeric_limits_and_mixed_full_scan(self) -> None:
        with redirect_stderr(io.StringIO()), self.assertRaises(SystemExit) as raised:
            adapter._parser().parse_args(["--max-messages", "0"])
        self.assertEqual(2, raised.exception.code)

        with redirect_stderr(io.StringIO()), self.assertRaises(SystemExit) as raised:
            adapter.main(["--full-scan", "--max-messages", "10"])
        self.assertEqual(2, raised.exception.code)


if __name__ == "__main__":
    unittest.main()
