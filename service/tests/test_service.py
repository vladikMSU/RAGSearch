from __future__ import annotations

import json
import math
import os
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
from dataclasses import replace
from pathlib import Path
from unittest.mock import patch

from ragsearch_service.app import SearchService
from ragsearch_service.config import Settings
from ragsearch_service.errors import ValidationError
from ragsearch_service.http_api import create_http_server


class ControlledEmbedder:
    name = "controlled-test-v1"
    dimensions = 2
    fingerprint = "sha256:" + "1" * 64

    @staticmethod
    def _unit(cosine: float) -> list[float]:
        return [cosine, math.sqrt(1.0 - cosine * cosine)]

    def embed_many(self, texts: list[str]) -> list[list[float]]:
        vectors: list[list[float]] = []
        for text in texts:
            if text in {"controlled-query", "controlled query"} or "near-vector" in text:
                vectors.append(self._unit(1.0))
            elif "close-vector" in text:
                vectors.append(self._unit(0.96))
            elif "window-vector" in text:
                vectors.append(self._unit(0.94))
            elif "far-vector" in text:
                vectors.append(self._unit(0.70))
            elif "very-weak-vector" in text:
                vectors.append(self._unit(0.20))
            elif "weak-vector" in text:
                vectors.append(self._unit(0.35))
            else:
                vectors.append(self._unit(0.0))
        return vectors


class AdversarialShortQueryEmbedder:
    name = "adversarial-short-query-v1"
    dimensions = 2
    fingerprint = "sha256:" + "2" * 64

    @staticmethod
    def _unit(cosine: float) -> list[float]:
        return [cosine, math.sqrt(1.0 - cosine * cosine)]

    def embed_many(self, texts: list[str]) -> list[list[float]]:
        vectors: list[list[float]] = []
        for text in texts:
            folded = text.casefold()
            if folded in {"кибер", "спорт", "киберспорт"}:
                vectors.append(self._unit(1.0))
            elif "киберспорт" in folded:
                vectors.append(self._unit(0.20))
            elif "куку" in folded:
                vectors.append(self._unit(1.0))
            else:
                vectors.append(self._unit(0.0))
        return vectors


class SettingsTests(unittest.TestCase):
    def test_default_path_requires_localappdata(self) -> None:
        for environment in ({}, {"LOCALAPPDATA": ""}):
            with self.subTest(environment=environment):
                with patch.dict(os.environ, environment, clear=True):
                    with self.assertRaisesRegex(RuntimeError, "LOCALAPPDATA is required"):
                        Settings.default()


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
            "body": "Релиз проекта перенесли на октябрь. OriginalUniqueTerm присутствует здесь.",
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
        self.assertEqual(3, stats["schema_version"])
        self.assertEqual("hashing-v1-256", stats["embedding_model"])
        self.assertEqual(256, stats["embedding_dim"])
        self.assertRegex(
            stats["embedding_fingerprint"],
            r"\Asha256:[0-9a-f]{64}\Z",
        )

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
                'SELECT COUNT(*) FROM chunks_fts WHERE chunks_fts MATCH "OriginalUniqueTerm"'
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

    def test_noncanonical_recipients_and_timestamps_are_rejected(self) -> None:
        outcome = self.service.ingest_messages(
            {"messages": [self.message(to=["user@example.test"])]}
        )

        self.assertEqual(0, outcome["accepted"])
        self.assertEqual(1, outcome["failed"])
        self.assertIn("to must be a string or null", outcome["errors"][0]["error"])

        outcome = self.service.ingest_messages(
            {"messages": [self.message(received_at="2026-08-11T09:00:00")]}
        )
        self.assertEqual(0, outcome["accepted"])
        self.assertEqual(1, outcome["failed"])
        self.assertIn("UTC offset", outcome["errors"][0]["error"])

        for value in ("", "2026-08-11 09:00:00Z", "not-a-timestamp"):
            with self.subTest(timestamp=value):
                outcome = self.service.ingest_messages(
                    {"messages": [self.message(received_at=value)]}
                )
                self.assertEqual(0, outcome["accepted"])
                self.assertEqual(1, outcome["failed"])
                self.assertIn("UTC offset", outcome["errors"][0]["error"])

    def test_timestamps_are_canonical_utc_and_filters_compare_instants(self) -> None:
        messages = [
            self.message(
                entry_id="before",
                internet_message_id="<before@example.test>",
                received_at="2026-08-11T12:00:00+03:00",
                body="TemporalBoundaryMarker before",
            ),
            self.message(
                entry_id="boundary",
                internet_message_id="<boundary@example.test>",
                received_at="2026-08-11T05:00:00-05:00",
                body="TemporalBoundaryMarker boundary",
            ),
            self.message(
                entry_id="after",
                internet_message_id="<after@example.test>",
                received_at="2026-08-11T10:00:00.000001Z",
                body="TemporalBoundaryMarker after",
            ),
        ]
        self.assertEqual(
            3,
            self.service.ingest_messages({"messages": messages})["accepted"],
        )

        with self.service.database.session() as connection:
            stored = {
                row["entry_id"]: row["received_at"]
                for row in connection.execute(
                    "SELECT entry_id, received_at FROM messages"
                )
            }
        self.assertEqual("2026-08-11T09:00:00.000000Z", stored["before"])
        self.assertEqual("2026-08-11T10:00:00.000000Z", stored["boundary"])
        self.assertEqual("2026-08-11T10:00:00.000001Z", stored["after"])

        response = self.service.search(
            {
                "query": "TemporalBoundaryMarker",
                "limit": 10,
                "filters": {
                    "received_from": "2026-08-11T12:00:00+02:00",
                    "received_to": "2026-08-11T10:00:00Z",
                },
            }
        )
        self.assertEqual(
            ["boundary"],
            [result["entry_id"] for result in response["results"]],
        )

        for value in (None, "", "2026-08-11T10:00:00", "not-a-timestamp"):
            with self.subTest(filter_timestamp=value):
                with self.assertRaisesRegex(ValidationError, "UTC offset"):
                    self.service.search(
                        {
                            "query": "TemporalBoundaryMarker",
                            "filters": {"received_from": value},
                        }
                    )

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

    def test_search_uses_adaptive_vector_cutoff_distance_order_and_top_n(self) -> None:
        settings = replace(
            Settings.explicit(Path(self.temporary.name) / "controlled-ranking"),
            search_result_limit=2,
        )
        service = SearchService(settings, embedder=ControlledEmbedder())
        messages = [
            self.message(
                entry_id=f"entry-{label}",
                internet_message_id=f"<{label}@example.test>",
                subject=label,
                body=label,
            )
            for label in ("far-vector", "window-vector", "near-vector", "close-vector")
        ]
        self.assertEqual(
            4,
            service.ingest_messages({"messages": messages})["accepted"],
        )

        response = service.search({"query": "controlled query", "limit": 100})

        self.assertEqual(
            "lexical_gate_then_vector_distance_asc",
            response["ranking"],
        )
        self.assertEqual(4, response["candidate_count"])
        self.assertEqual(3, response["eligible_count"])
        self.assertEqual(2, response["max_results"])
        self.assertAlmostEqual(0.9, response["cutoff_similarity"], places=6)
        self.assertAlmostEqual(0.1, response["cutoff_distance"], places=6)
        self.assertEqual(
            ["entry-near-vector", "entry-close-vector"],
            [item["entry_id"] for item in response["results"]],
        )
        self.assertEqual([1, 2], [item["rank"] for item in response["results"]])
        self.assertLessEqual(
            response["results"][0]["vector_distance"],
            response["results"][1]["vector_distance"],
        )
        self.assertIn("hybrid_score", response["results"][0])

    def test_search_returns_empty_when_best_vector_is_below_floor(self) -> None:
        settings = Settings.explicit(Path(self.temporary.name) / "controlled-cutoff")
        service = SearchService(settings, embedder=ControlledEmbedder())
        messages = [
            self.message(
                entry_id=f"entry-{label}",
                internet_message_id=f"<{label}@example.test>",
                subject=label,
                body=label,
            )
            for label in ("weak-vector", "very-weak-vector")
        ]
        service.ingest_messages({"messages": messages})

        response = service.search({"query": "controlled query", "limit": 100})

        self.assertEqual(2, response["candidate_count"])
        self.assertEqual(0, response["eligible_count"])
        self.assertEqual([], response["results"])
        self.assertAlmostEqual(0.4, response["cutoff_similarity"], places=6)
        self.assertAlmostEqual(0.6, response["cutoff_distance"], places=6)

    def test_existing_index_rejects_a_different_embedding_model(self) -> None:
        self.service.ingest_messages(
            {"messages": [self.message(body="SingleModelContractMarker")]}
        )
        with self.assertRaisesRegex(RuntimeError, "embedding contract"):
            SearchService(self.settings, embedder=ControlledEmbedder())
        self.assertEqual(1, self.service.stats()["messages"])

    def test_existing_index_rejects_same_model_label_with_new_fingerprint(self) -> None:
        class ChangedImplementation(ControlledEmbedder):
            fingerprint = "sha256:" + "3" * 64

        settings = Settings.explicit(Path(self.temporary.name) / "fingerprint-change")
        service = SearchService(settings, embedder=ControlledEmbedder())
        service.ingest_messages(
            {"messages": [self.message(body="FingerprintContractMarker")]}
        )

        with self.assertRaisesRegex(RuntimeError, "embedding contract"):
            SearchService(settings, embedder=ChangedImplementation())
        self.assertEqual(1, service.stats()["messages"])

    def test_clear_index_allows_an_explicit_embedding_model_change(self) -> None:
        settings = Settings.explicit(Path(self.temporary.name) / "model-change")
        initial_service = SearchService(settings)
        initial_service.ingest_messages({"messages": [self.message()]})
        initial_service.clear_index()

        switched_service = SearchService(settings, embedder=ControlledEmbedder())
        switched_service.ingest_messages(
            {"messages": [self.message(body="controlled-query")]}
        )

        response = switched_service.search({"query": "controlled-query", "limit": 2})
        self.assertEqual("entry-1", response["results"][0]["entry_id"])
        self.assertEqual(
            "controlled-test-v1",
            switched_service.stats()["embedding_model"],
        )
        self.assertEqual(2, switched_service.stats()["embedding_dim"])

    def test_vector_candidate_pool_keeps_best_chunk_per_message(self) -> None:
        settings = replace(
            Settings.explicit(Path(self.temporary.name) / "message-diversity"),
            chunk_chars=64,
            chunk_overlap_chars=0,
        )
        service = SearchService(settings, embedder=ControlledEmbedder())
        service.ingest_messages(
            {
                "messages": [
                    self.message(
                        entry_id="many-near-chunks",
                        internet_message_id="<many-near-chunks@example.test>",
                        subject="Many chunks",
                        body=" ".join(["near-vector"] * 40),
                    ),
                    self.message(
                        entry_id="one-close-chunk",
                        internet_message_id="<one-close-chunk@example.test>",
                        subject="One chunk",
                        body="close-vector",
                    ),
                ]
            }
        )

        results = service.database.search(
            "controlled-query",
            limit=10,
            filters={},
            embedder=service.embedder,
            candidate_limit=2,
        )

        self.assertEqual(
            {"many-near-chunks", "one-close-chunk"},
            {item["entry_id"] for item in results},
        )

    def test_single_word_prefix_and_suffix_literal_hits_beat_dense_distractor(self) -> None:
        settings = Settings.explicit(Path(self.temporary.name) / "trigram-ranking")
        service = SearchService(settings, embedder=AdversarialShortQueryEmbedder())
        service.ingest_messages(
            {
                "messages": [
                    self.message(
                        entry_id="target-cybersport",
                        internet_message_id="<target-cybersport@example.test>",
                        subject="Target",
                        body="КИБЕРСПОРТ и длинная история переписки",
                    ),
                    self.message(
                        entry_id="dense-distractor",
                        internet_message_id="<dense-distractor@example.test>",
                        subject="Distractor",
                        body="куку",
                    ),
                ]
            }
        )

        expected_kinds = {
            "киберспорт": "token",
            "кибер": "prefix",
            "спорт": "substring",
        }
        for query, expected_kind in expected_kinds.items():
            with self.subTest(query=query):
                response = service.search({"query": query, "limit": 25})
                self.assertTrue(response["lexical_gate"])
                self.assertEqual(1, response["lexical_match_count"])
                self.assertEqual(
                    ["target-cybersport"],
                    [item["entry_id"] for item in response["results"]],
                )
                self.assertEqual(
                    expected_kind,
                    response["results"][0]["lexical_match_kind"],
                )
                self.assertEqual(
                    "lexical_" + expected_kind,
                    response["results"][0]["ranking_basis"],
                )

    def test_current_model_literal_hit_survives_full_vector_pool(self) -> None:
        settings = replace(
            Settings.explicit(Path(self.temporary.name) / "literal-vector-pool"),
            search_candidate_message_limit=2,
        )
        service = SearchService(settings, embedder=AdversarialShortQueryEmbedder())
        messages = [
            self.message(
                entry_id="target-cybersport",
                internet_message_id="<target-cybersport@example.test>",
                subject="Target",
                body="КИБЕРСПОРТ и длинная история переписки",
            )
        ]
        messages.extend(
            self.message(
                entry_id=f"dense-distractor-{index}",
                internet_message_id=f"<dense-distractor-{index}@example.test>",
                subject=f"Distractor {index}",
                body="куку",
            )
            for index in range(3)
        )
        self.assertEqual(
            4,
            service.ingest_messages({"messages": messages})["accepted"],
        )

        response = service.search({"query": "кибер", "limit": 2})

        self.assertTrue(response["lexical_gate"])
        self.assertEqual(1, response["lexical_match_count"])
        self.assertEqual(
            ["target-cybersport"],
            [item["entry_id"] for item in response["results"]],
        )
        self.assertEqual("lexical_prefix", response["results"][0]["ranking_basis"])

    def test_schema_v2_is_rejected_without_modifying_it(self) -> None:
        settings = Settings.explicit(Path(self.temporary.name) / "schema-v2")
        service = SearchService(settings)
        with service.database.session() as connection:
            for trigger in (
                "chunks_trigram_ai",
                "chunks_trigram_ad",
                "chunks_trigram_au",
            ):
                connection.execute(f"DROP TRIGGER {trigger}")
            connection.execute("DROP TABLE chunks_trigram")
            connection.execute("PRAGMA user_version = 2")

        with self.assertRaisesRegex(
            RuntimeError,
            "Unsupported database schema version 2; expected 3",
        ):
            service.database.initialize()

        with service.database.session() as connection:
            schema_version = int(
                connection.execute("PRAGMA user_version").fetchone()[0]
            )
            self.assertEqual(
                0,
                connection.execute(
                    "SELECT COUNT(*) FROM sqlite_master "
                    "WHERE type = 'table' AND name = 'chunks_trigram'"
                ).fetchone()[0],
            )
        self.assertEqual(2, schema_version)

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
            "to": "",
            "cc": "",
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
        self.assertEqual(
            "lexical_gate_then_vector_distance_asc",
            payload["ranking"],
        )
        self.assertEqual(5, payload["max_results"])
        self.assertEqual(1, payload["results"][0]["rank"])
        self.assertNotIn("score", payload["results"][0])
        self.assertIn("vector_similarity", payload["results"][0])
        self.assertIn("vector_distance", payload["results"][0])

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
        self.assertIsNotNone(before["embedding_fingerprint"])
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
        self.assertIsNone(after["embedding_model"])
        self.assertIsNone(after["embedding_dim"])
        self.assertIsNone(after["embedding_fingerprint"])
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
