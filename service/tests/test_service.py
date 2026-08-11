from __future__ import annotations

import json
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
from pathlib import Path

from ragsearch_service.app import SearchService
from ragsearch_service.config import Settings
from ragsearch_service.http_api import create_http_server


class SearchServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.settings = Settings.explicit(Path(self.temporary.name))
        self.service = SearchService(self.settings)

    def message(self, **changes: object) -> dict[str, object]:
        payload: dict[str, object] = {
            "entry_id": "entry-1",
            "store_id": "store-1",
            "folder_entry_id": "folder-1",
            "folder_path": "\\Mailbox\\Inbox",
            "store_name": "Mailbox",
            "subject": "План запуска продукта",
            "sender_name": "Алексей",
            "sender_email": "alex@example.test",
            "to": "user@example.test",
            "cc": "",
            "sent_at": "2026-08-11T08:59:00Z",
            "received_at": "2026-08-11T09:00:00Z",
            "modified_at": "2026-08-11T09:01:00Z",
            "internet_message_id": "<entry-1@example.test>",
            "conversation_id": "conversation-1",
            "body": "Релиз проекта перенесли на октябрь. LegacyUniqueTerm присутствует здесь.",
            "attachments": [],
        }
        payload.update(changes)
        return payload

    def test_ingest_attachment_search_preserves_outlook_identity(self) -> None:
        attachment_path = self.settings.spool_dir / "notes.txt"
        attachment_path.write_text(
            "Космический трактор используется как внутреннее название проекта.",
            encoding="utf-8",
        )
        message = self.message(
            attachments=[
                {
                    "name": "notes.txt",
                    "size": attachment_path.stat().st_size,
                    "content_type": "text/plain",
                    "temp_path": str(attachment_path),
                }
            ]
        )

        outcome = self.service.ingest_messages({"messages": [message]})
        self.assertEqual({"accepted": 1, "failed": 0, "errors": []}, outcome)

        results = self.service.search({"query": "космический трактор", "limit": 10})[
            "results"
        ]
        self.assertTrue(results)
        result = results[0]
        self.assertEqual("entry-1", result["entry_id"])
        self.assertEqual("store-1", result["store_id"])
        self.assertEqual("folder-1", result["folder_entry_id"])
        self.assertEqual("Mailbox", result["store_name"])
        self.assertEqual("<entry-1@example.test>", result["internet_message_id"])
        self.assertEqual("conversation-1", result["conversation_id"])
        self.assertIn("attachment:notes.txt", result["matched_sources"])

        stats = self.service.stats()
        self.assertEqual(1, stats["messages"])
        self.assertEqual(1, stats["attachments"])
        self.assertGreaterEqual(stats["chunks"], 3)
        self.assertEqual({"extracted": 1}, stats["attachment_extraction"])

    def test_upsert_replaces_only_that_messages_children(self) -> None:
        self.assertEqual(
            1,
            self.service.ingest_messages({"messages": [self.message()]})["accepted"],
        )
        updated = self.message(
            folder_entry_id="folder-2",
            folder_path="\\Mailbox\\Archive",
            subject="Обновлённая тема",
            body="ReplacementUniqueTerm — новое содержимое.",
        )
        self.assertEqual(
            1,
            self.service.ingest_messages({"messages": [updated]})["accepted"],
        )

        stats = self.service.stats()
        self.assertEqual(1, stats["messages"])
        self.assertEqual(0, stats["attachments"])
        with self.service.database.session() as connection:
            stale = connection.execute(
                'SELECT COUNT(*) FROM chunks_fts WHERE chunks_fts MATCH "LegacyUniqueTerm"'
            ).fetchone()[0]
        self.assertEqual(0, stale)

        result = self.service.search({"query": "ReplacementUniqueTerm", "limit": 1})[
            "results"
        ][0]
        self.assertEqual("folder-2", result["folder_entry_id"])
        self.assertEqual("\\Mailbox\\Archive", result["folder_path"])

    def test_batch_reports_per_item_failures(self) -> None:
        invalid = self.message(entry_id="entry-bad")
        invalid.pop("store_id")
        outcome = self.service.ingest_messages(
            {"messages": [self.message(), invalid]}
        )
        self.assertEqual(1, outcome["accepted"])
        self.assertEqual(1, outcome["failed"])
        self.assertEqual(1, len(outcome["errors"]))
        self.assertEqual(1, outcome["errors"][0]["index"])

    def test_attachment_outside_spool_is_rejected(self) -> None:
        outside = Path(self.temporary.name) / "outside.txt"
        outside.write_text("must not be read", encoding="utf-8")
        message = self.message(
            attachments=[
                {
                    "name": "outside.txt",
                    "size": outside.stat().st_size,
                    "content_type": "text/plain",
                    "temp_path": str(outside),
                }
            ]
        )
        outcome = self.service.ingest_messages({"messages": [message]})
        self.assertEqual(0, outcome["accepted"])
        self.assertEqual(1, outcome["failed"])
        self.assertIn("outside", outcome["errors"][0]["error"].casefold())
        self.assertEqual(0, self.service.stats()["messages"])

    def test_search_filter_is_applied_to_lexical_and_vector_paths(self) -> None:
        first = self.message(body="Общий МаркерПроекта альфа")
        second = self.message(
            entry_id="entry-2",
            store_id="store-2",
            internet_message_id="<entry-2@example.test>",
            body="Общий МаркерПроекта бета",
        )
        self.service.ingest_messages({"messages": [first, second]})
        results = self.service.search(
            {
                "query": "МаркерПроекта",
                "limit": 10,
                "filters": {"store_id": "store-2"},
            }
        )["results"]
        self.assertTrue(results)
        self.assertEqual({"store-2"}, {item["store_id"] for item in results})

    def test_token_is_stable_for_explicit_test_path(self) -> None:
        original = self.service.token
        second_instance = SearchService(self.settings)
        self.assertEqual(original, second_instance.token)
        self.assertEqual(original, self.settings.token_path.read_text(encoding="ascii").strip())


class HTTPAPITests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        settings = Settings.explicit(Path(self.temporary.name), port=0)
        self.service = SearchService(settings)
        self.server = create_http_server(self.service)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.addCleanup(self._stop_server)
        self.base_url = f"http://127.0.0.1:{self.server.server_port}"

    def _stop_server(self) -> None:
        self.server.shutdown()
        self.server.server_close()
        self.thread.join(timeout=5)

    def _message(self, **changes: object) -> dict[str, object]:
        message: dict[str, object] = {
            "entry_id": "http-entry",
            "store_id": "http-store",
            "folder_entry_id": "http-folder",
            "folder_path": "\\HTTP\\Inbox",
            "store_name": "HTTP",
            "subject": "HTTP semantic test",
            "sender_name": "Test",
            "sender_email": "test@example.test",
            "to": [],
            "cc": [],
            "sent_at": None,
            "received_at": "2026-08-11T10:00:00Z",
            "modified_at": None,
            "internet_message_id": "",
            "conversation_id": "",
            "body": "Уникальный СинийДирижабль находится в письме.",
            "attachments": [],
        }
        message.update(changes)
        return message

    def _request(
        self,
        method: str,
        path: str,
        payload: object | None = None,
        *,
        token: bool = True,
    ) -> tuple[int, dict[str, object]]:
        data = None
        headers: dict[str, str] = {}
        if payload is not None:
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json"
        if token:
            headers["X-RAGSearch-Token"] = self.service.token
        request = urllib.request.Request(
            self.base_url + path,
            data=data,
            headers=headers,
            method=method,
        )
        try:
            response = urllib.request.urlopen(request, timeout=5)
        except urllib.error.HTTPError as error:
            return error.code, json.loads(error.read().decode("utf-8"))
        with response:
            return response.status, json.loads(response.read().decode("utf-8"))

    def test_health_is_minimal_and_v1_requires_token(self) -> None:
        status, payload = self._request("GET", "/health", token=False)
        self.assertEqual(200, status)
        self.assertEqual({"status": "ok"}, payload)

        status, payload = self._request("GET", "/v1/stats", token=False)
        self.assertEqual(401, status)
        self.assertEqual("unauthorized", payload["error"])

    def test_http_ingest_and_search_contract(self) -> None:
        message = self._message()
        status, outcome = self._request(
            "POST", "/v1/messages", {"messages": [message]}
        )
        self.assertEqual(200, status)
        self.assertEqual(1, outcome["accepted"])

        status, payload = self._request(
            "POST", "/v1/search", {"query": "СинийДирижабль", "limit": 5}
        )
        self.assertEqual(200, status)
        self.assertEqual("http-entry", payload["results"][0]["entry_id"])
        self.assertEqual("http-store", payload["results"][0]["store_id"])
        self.assertEqual("http-folder", payload["results"][0]["folder_entry_id"])
        self.assertEqual("HTTP", payload["results"][0]["store_name"])
        self.assertEqual("", payload["results"][0]["internet_message_id"])
        self.assertEqual("", payload["results"][0]["conversation_id"])

    def test_clear_index_requires_token(self) -> None:
        self.service.ingest_messages({"messages": [self._message()]})

        status, payload = self._request("DELETE", "/v1/index", token=False)

        self.assertEqual(401, status)
        self.assertEqual("unauthorized", payload["error"])
        self.assertEqual(1, self.service.stats()["messages"])

    def test_clear_index_is_complete_idempotent_and_service_remains_usable(self) -> None:
        attachment_path = self.service.settings.spool_dir / "reset-test.txt"
        attachment_path.write_text("ResetAttachmentUniqueTerm", encoding="utf-8")
        message = self._message(
            body="ResetBodyUniqueTerm",
            attachments=[
                {
                    "name": attachment_path.name,
                    "size": attachment_path.stat().st_size,
                    "content_type": "text/plain",
                    "temp_path": str(attachment_path),
                }
            ],
        )
        self.service.ingest_messages({"messages": [message]})
        before = self.service.stats()
        original_token = self.service.token

        status, payload = self._request("DELETE", "/v1/index")

        self.assertEqual(200, status)
        self.assertEqual(before["messages"], payload["deleted_messages"])
        self.assertEqual(before["attachments"], payload["deleted_attachments"])
        self.assertEqual(before["chunks"], payload["deleted_chunks"])
        after = self.service.stats()
        self.assertEqual(0, after["messages"])
        self.assertEqual(0, after["attachments"])
        self.assertEqual(0, after["chunks"])
        with self.service.database.session() as connection:
            fts_matches = connection.execute(
                "SELECT COUNT(*) FROM chunks_fts WHERE chunks_fts MATCH ?",
                ("ResetAttachmentUniqueTerm",),
            ).fetchone()[0]
        self.assertEqual(0, fts_matches)
        self.assertTrue(self.service.settings.database_path.is_file())
        self.assertTrue(attachment_path.is_file())
        self.assertEqual(original_token, self.service.token)
        self.assertEqual(
            original_token,
            self.service.settings.token_path.read_text(encoding="ascii").strip(),
        )

        status, payload = self._request("DELETE", "/v1/index")
        self.assertEqual(200, status)
        self.assertEqual(
            {
                "deleted_messages": 0,
                "deleted_attachments": 0,
                "deleted_chunks": 0,
            },
            payload,
        )

        status, outcome = self._request(
            "POST", "/v1/messages", {"messages": [message]}
        )
        self.assertEqual(200, status)
        self.assertEqual(1, outcome["accepted"])
        status, search = self._request(
            "POST", "/v1/search", {"query": "ResetAttachmentUniqueTerm", "limit": 5}
        )
        self.assertEqual(200, status)
        self.assertEqual("http-entry", search["results"][0]["entry_id"])


if __name__ == "__main__":
    unittest.main()
