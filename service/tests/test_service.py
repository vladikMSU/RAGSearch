from __future__ import annotations

import base64
import contextlib
import io
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

from ragsearch_service.__main__ import _parser
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


def document(**changes: object) -> dict[str, object]:
    body = "Релиз проекта перенесли на октябрь. OriginalUniqueTerm присутствует здесь."
    payload: dict[str, object] = {
        "source_key": "test:document-1",
        "kind": "email",
        "title": "План запуска продукта",
        "metadata": {
            "sender_name": "Алексей",
            "sender_email": "alex@example.test",
            "received_at": "2026-08-11T09:00:00Z",
            "folder_path": "\\Mailbox\\Inbox",
        },
        "locator": {
            "connector": "example",
            "resource": {"collection": "alpha", "item": 7},
            "flags": ["opaque", True],
        },
        "parts": [
            {
                "key": "body",
                "kind": "body",
                "name": "body",
                "media_type": "text/plain",
                "size": len(body.encode("utf-8")),
                "text": body,
                "truncated": False,
            }
        ],
    }
    payload.update(changes)
    return payload


def text_document(source_key: str, label: str) -> dict[str, object]:
    return document(
        source_key=source_key,
        title=label,
        metadata={"label": label},
        locator={"connector": "test", "key": source_key},
        parts=[
            {
                "key": "body",
                "kind": "body",
                "name": "body",
                "media_type": "text/plain",
                "size": len(label.encode("utf-8")),
                "text": label,
            }
        ],
    )


class SettingsTests(unittest.TestCase):
    def test_default_path_requires_localappdata(self) -> None:
        for environment in ({}, {"LOCALAPPDATA": ""}):
            with self.subTest(environment=environment):
                with patch.dict(os.environ, environment, clear=True):
                    with self.assertRaisesRegex(RuntimeError, "LOCALAPPDATA is required"):
                        Settings.default()

    def test_service_configuration_has_no_spool_contract(self) -> None:
        settings = Settings.explicit(Path("isolated"))
        self.assertFalse(hasattr(settings, "spool_dir"))
        self.assertFalse(hasattr(settings, "delete_spool_after_ingest"))
        self.assertEqual(48 * 1024 * 1024, settings.request_limit_bytes)
        self.assertEqual(8 * 1024 * 1024, settings.inline_part_limit_bytes)
        self.assertEqual(8 * 1024 * 1024, settings.extracted_text_limit_chars)
        self.assertEqual(1 * 1024 * 1024, settings.metadata_json_limit_bytes)
        self.assertEqual(64 * 1024, settings.locator_json_limit_bytes)
        self.assertEqual(16, settings.opaque_json_max_depth)
        self.assertEqual(10_000, settings.opaque_json_max_nodes)
        self.assertEqual(16 * 1024 * 1024, settings.document_text_limit_chars)
        self.assertEqual(4 * 1024 * 1024, settings.metadata_search_text_limit_chars)
        self.assertEqual(8 * 1024 * 1024, settings.binary_extracted_text_limit_chars)
        self.assertEqual(16_384, settings.document_chunk_limit)
        self.assertEqual(8_192, settings.query_limit_chars)
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            _parser().parse_args(["--spool-dir", "somewhere"])


class SearchServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.settings = Settings.explicit(Path(self.temporary.name))
        self.service = SearchService(self.settings)

    def test_inline_attachment_search_and_opaque_locator_round_trip(self) -> None:
        attachment = "Космический трактор используется как внутреннее название проекта.".encode(
            "utf-8"
        )
        locator = {
            "connector": "filesystem-v9",
            "path": ["vault", "note.md"],
            "revision": {"generation": 42, "current": True},
        }
        payload = document(
            locator=locator,
            parts=[
                document()["parts"][0],
                {
                    "key": "attachment:0",
                    "kind": "attachment",
                    "name": "notes.txt",
                    "media_type": "text/plain",
                    "size": len(attachment),
                    "content_base64": base64.b64encode(attachment).decode("ascii"),
                },
            ],
        )

        self.assertEqual(
            {"source_key": "test:document-1", "status": "upserted"},
            self.service.ingest_document(payload),
        )
        result = self.service.search({"query": "космический трактор", "limit": 10})[
            "results"
        ][0]
        self.assertEqual("test:document-1", result["source_key"])
        self.assertEqual("email", result["kind"])
        self.assertEqual("План запуска продукта", result["title"])
        self.assertEqual(payload["metadata"], result["metadata"])
        self.assertEqual(locator, result["locator"])
        self.assertEqual("attachment:notes.txt", result["snippet_part"])
        self.assertIn("attachment:notes.txt", result["matched_parts"])

        stats = self.service.stats()
        self.assertEqual(1, stats["documents"])
        self.assertEqual(2, stats["parts"])
        self.assertGreaterEqual(stats["chunks"], 3)
        self.assertEqual(
            {"extracted": 1, "provided_text": 1},
            stats["part_extraction"],
        )
        self.assertEqual(4, stats["schema_version"])

    def test_metadata_only_part_name_remains_searchable(self) -> None:
        payload = document(
            parts=[
                document()["parts"][0],
                {
                    "key": "attachment:0",
                    "kind": "attachment",
                    "name": "QuarterlyRoadmap.xlsx",
                    "media_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "size": 42,
                },
            ]
        )
        self.service.ingest_document(payload)

        result = self.service.search({"query": "QuarterlyRoadmap", "limit": 1})[
            "results"
        ][0]
        self.assertEqual("test:document-1", result["source_key"])
        self.assertEqual("metadata", result["snippet_part"])
        self.assertIn("metadata", result["matched_parts"])

    def test_opaque_json_is_bounded_by_bytes_depth_and_nodes(self) -> None:
        settings = replace(
            Settings.explicit(Path(self.temporary.name) / "bounded-json"),
            metadata_json_limit_bytes=48,
            locator_json_limit_bytes=32,
            opaque_json_max_depth=2,
            opaque_json_max_nodes=5,
        )
        service = SearchService(settings)
        cases = [
            (
                document(metadata={"value": "я" * 30}, locator={}),
                "metadata exceeds 48 UTF-8 JSON bytes",
            ),
            (
                document(metadata={}, locator={"value": "x" * 40}),
                "locator exceeds 32 UTF-8 JSON bytes",
            ),
            (
                document(metadata={"a": {"b": {"c": 1}}}, locator={}),
                "metadata exceeds JSON depth 2",
            ),
            (
                document(metadata={"values": [1, 2, 3, 4]}, locator={}),
                "metadata exceeds 5 JSON nodes",
            ),
            (
                document(metadata={"nested": {"__type": "evil:#x"}}, locator={}),
                "metadata contains reserved JSON key __type",
            ),
        ]
        for payload, message in cases:
            with self.subTest(message=message), self.assertRaisesRegex(
                ValidationError, message
            ):
                service.ingest_document(payload)
        self.assertEqual(0, service.stats()["documents"])

    def test_opaque_json_numbers_must_fit_the_host_finite_double_range(self) -> None:
        accepted = document(
            source_key="test:large-finite-number",
            metadata={"large": 10**100, "finite_float": 1.7976931348623157e308},
            locator={"negative": -(10**100)},
        )
        self.service.ingest_document(accepted)
        result = self.service.search({"query": "large", "limit": 1})["results"][0]
        self.assertEqual(10**100, result["metadata"]["large"])

        cases = [
            document(
                source_key="test:overflow-metadata-int",
                metadata={"number": 10**1000},
            ),
            document(
                source_key="test:overflow-locator-int",
                locator={"number": -(10**1000)},
            ),
            document(source_key="test:nan", metadata={"number": float("nan")}),
            document(source_key="test:positive-inf", metadata={"number": float("inf")}),
            document(source_key="test:negative-inf", locator={"number": float("-inf")}),
        ]
        for payload in cases:
            with self.subTest(source_key=payload["source_key"]), self.assertRaisesRegex(
                ValidationError,
                "finite IEEE-754 range",
            ):
                self.service.ingest_document(payload)

    def test_source_key_upsert_replaces_only_that_documents_children(self) -> None:
        self.service.ingest_document(document())
        updated = document(
            title="Обновлённый документ",
            metadata={"folder": "archive"},
            locator={"connector": "new-source", "id": "replacement"},
            parts=[
                {
                    "key": "replacement",
                    "kind": "body",
                    "name": "replacement",
                    "media_type": "text/plain",
                    "size": 0,
                    "text": "ReplacementUniqueTerm — новое содержимое.",
                }
            ],
        )
        self.service.ingest_document(updated)

        stats = self.service.stats()
        self.assertEqual(1, stats["documents"])
        self.assertEqual(1, stats["parts"])
        with self.service.database.session() as connection:
            stale = connection.execute(
                'SELECT COUNT(*) FROM chunks_fts WHERE chunks_fts MATCH "OriginalUniqueTerm"'
            ).fetchone()[0]
        self.assertEqual(0, stale)
        result = self.service.search({"query": "ReplacementUniqueTerm", "limit": 1})[
            "results"
        ][0]
        self.assertEqual("test:document-1", result["source_key"])
        self.assertEqual(updated["metadata"], result["metadata"])
        self.assertEqual(updated["locator"], result["locator"])

    def test_source_key_is_opaque_and_preserves_whitespace(self) -> None:
        self.service.ingest_document(
            text_document("test:identity", "first identity body")
        )
        self.service.ingest_document(
            text_document(" test:identity ", "second identity body")
        )

        self.assertEqual(2, self.service.stats()["documents"])
        keys = {
            row["source_key"]
            for row in self.service.search(
                {"query": "identity body", "limit": 10}
            )["results"]
        }
        self.assertEqual({"test:identity", " test:identity "}, keys)

    def test_document_contract_is_single_strict_object(self) -> None:
        with self.assertRaisesRegex(ValidationError, "Unknown document fields"):
            self.service.ingest_document({"documents": [document()]})
        with self.assertRaisesRegex(ValidationError, "Unknown document fields"):
            self.service.ingest_document(document(entry_id="outlook-shaped"))
        with self.assertRaisesRegex(ValidationError, "Duplicate part key"):
            part = document()["parts"][0]
            self.service.ingest_document(document(parts=[part, part]))
        self.service.ingest_document(
            document(source_key="test:max-title", title="t" * 65_536)
        )
        with self.assertRaisesRegex(ValidationError, "title exceeds 65536 characters"):
            self.service.ingest_document(
                document(source_key="test:oversize-title", title="t" * 65_537)
            )
        self.assertEqual(1, self.service.stats()["documents"])

    def test_unpaired_surrogates_are_rejected_in_all_neutral_string_locations(self) -> None:
        cases = [
            document(source_key="bad\ud800source"),
            document(kind="bad\udfffkind"),
            document(title="bad\ud800title"),
            document(metadata={"nested": {"bad\ud800key": "value"}}),
            document(metadata={"nested": ["bad\udfffvalue"]}),
            document(locator={"nested": {"value": "bad\ud800locator"}}),
            document(
                parts=[
                    {
                        "key": "bad\ud800key",
                        "kind": "body",
                        "name": "bad\udfffname",
                        "media_type": "text/plain",
                        "size": 1,
                        "text": "x",
                    }
                ]
            ),
            document(
                parts=[
                    {
                        "key": "body",
                        "kind": "body",
                        "size": 1,
                        "text": "bad\ud800text",
                    }
                ]
            ),
        ]
        for payload in cases:
            with self.subTest(payload=ascii(payload)), self.assertRaisesRegex(
                ValidationError,
                "unpaired UTF-16 surrogate",
            ):
                self.service.ingest_document(payload)
        with self.assertRaisesRegex(ValidationError, "unpaired UTF-16 surrogate"):
            self.service.search({"query": "bad\udfffquery", "limit": 1})
        self.assertEqual(0, self.service.stats()["documents"])

    def test_valid_astral_unicode_remains_accepted_and_round_trips(self) -> None:
        payload = document(
            source_key="test:astral:😀",
            kind="note-🚀",
            title="Launch 😀",
            metadata={"emoji😀": {"value": "🚀"}},
            locator={"connector": "test", "resource": "🌍"},
            parts=[
                {
                    "key": "body:😀",
                    "kind": "body",
                    "name": "Body 🚀",
                    "media_type": "text/plain",
                    "size": len("Valid 🌍 text".encode("utf-8")),
                    "text": "Valid 🌍 text",
                }
            ],
        )
        self.service.ingest_document(payload)
        with self.service.database.session() as connection:
            row = connection.execute(
                "SELECT source_key, kind, title, metadata_json, locator_json "
                "FROM documents WHERE source_key = ?",
                (payload["source_key"],),
            ).fetchone()
        self.assertEqual(payload["source_key"], row["source_key"])
        self.assertEqual(payload["kind"], row["kind"])
        self.assertEqual(payload["title"], row["title"])
        self.assertEqual(payload["metadata"], json.loads(row["metadata_json"]))
        self.assertEqual(payload["locator"], json.loads(row["locator_json"]))
        self.assertEqual([], self.service.search({"query": "😀", "limit": 1})["results"])

    def test_invalid_base64_size_mismatch_and_oversize_are_rejected(self) -> None:
        cases = [
            (
                document(
                    parts=[
                        {
                            "key": "bad-base64",
                            "kind": "attachment",
                            "size": 3,
                            "content_base64": "not base64!",
                        }
                    ]
                ),
                "not valid base64",
            ),
            (
                document(
                    parts=[
                        {
                            "key": "bad-size",
                            "kind": "attachment",
                            "size": 999,
                            "content_base64": base64.b64encode(b"abc").decode("ascii"),
                        }
                    ]
                ),
                "does not match",
            ),
        ]
        for payload, error in cases:
            with self.subTest(error=error), self.assertRaisesRegex(ValidationError, error):
                self.service.ingest_document(payload)

        small_settings = replace(
            Settings.explicit(Path(self.temporary.name) / "small-inline"),
            inline_part_limit_bytes=4,
        )
        small_service = SearchService(small_settings)
        content = b"12345"
        with self.assertRaisesRegex(ValidationError, "inline part limit"):
            small_service.ingest_document(
                document(
                    parts=[
                        {
                            "key": "oversize",
                            "kind": "attachment",
                            "size": len(content),
                            "content_base64": base64.b64encode(content).decode("ascii"),
                        }
                    ]
                )
            )

    def test_text_and_base64_are_mutually_exclusive(self) -> None:
        with self.assertRaisesRegex(ValidationError, "both text and content_base64"):
            self.service.ingest_document(
                document(
                    parts=[
                        {
                            "key": "ambiguous",
                            "kind": "body",
                            "size": 1,
                            "text": "x",
                            "content_base64": "eA==",
                        }
                    ]
                )
            )

    def test_extracted_text_is_bounded_without_truncating_provided_text(self) -> None:
        settings = replace(
            Settings.explicit(Path(self.temporary.name) / "bounded-extraction"),
            extracted_text_limit_chars=4,
        )
        service = SearchService(settings)
        content = b"abcdefgh"
        service.ingest_document(
            document(
                source_key="test:extracted",
                metadata={},
                locator={},
                parts=[
                    {
                        "key": "attachment",
                        "kind": "attachment",
                        "name": "sample.txt",
                        "media_type": "text/plain",
                        "size": len(content),
                        "content_base64": base64.b64encode(content).decode("ascii"),
                    }
                ],
            )
        )
        provided = "provided-body-is-not-cropped"
        service.ingest_document(
            document(
                source_key="test:provided",
                metadata={},
                locator={},
                parts=[
                    {
                        "key": "body",
                        "kind": "body",
                        "name": "body",
                        "media_type": "text/plain",
                        "size": len(provided.encode("utf-8")),
                        "text": provided,
                    }
                ],
            )
        )

        with service.database.session() as connection:
            extracted_row = connection.execute(
                """
                SELECT p.extraction_status, p.extraction_error, c.text
                FROM parts p JOIN chunks c ON c.part_key = p.part_key
                    AND c.document_id = p.document_id
                JOIN documents d ON d.id = p.document_id
                WHERE d.source_key = 'test:extracted'
                """
            ).fetchone()
            provided_text = connection.execute(
                """
                SELECT c.text FROM chunks c
                JOIN documents d ON d.id = c.document_id
                WHERE d.source_key = 'test:provided' AND c.part_key = 'body'
                """
            ).fetchone()[0]
        self.assertEqual("extracted_truncated", extracted_row["extraction_status"])
        self.assertIn("4 characters", extracted_row["extraction_error"])
        self.assertEqual("abcd", extracted_row["text"])
        self.assertEqual(provided, provided_text)

    def test_search_rejects_removed_filters(self) -> None:
        self.service.ingest_document(document())
        with self.assertRaisesRegex(ValidationError, "Unknown search fields: filters"):
            self.service.search(
                {"query": "OriginalUniqueTerm", "filters": {"store_id": "store"}}
            )

    def test_search_uses_adaptive_vector_cutoff_distance_order_and_top_n(self) -> None:
        settings = replace(
            Settings.explicit(Path(self.temporary.name) / "controlled-ranking"),
            search_result_limit=2,
        )
        service = SearchService(settings, embedder=ControlledEmbedder())
        for label in ("far-vector", "window-vector", "near-vector", "close-vector"):
            service.ingest_document(text_document(f"test:{label}", label))

        response = service.search({"query": "controlled query", "limit": 100})
        self.assertEqual("lexical_gate_then_vector_distance_asc", response["ranking"])
        self.assertEqual(4, response["candidate_count"])
        self.assertEqual(3, response["eligible_count"])
        self.assertEqual(2, response["max_results"])
        self.assertAlmostEqual(0.9, response["cutoff_similarity"], places=6)
        self.assertEqual(
            ["test:near-vector", "test:close-vector"],
            [item["source_key"] for item in response["results"]],
        )
        self.assertEqual([1, 2], [item["rank"] for item in response["results"]])

    def test_search_returns_empty_when_best_vector_is_below_floor(self) -> None:
        settings = Settings.explicit(Path(self.temporary.name) / "controlled-cutoff")
        service = SearchService(settings, embedder=ControlledEmbedder())
        for label in ("weak-vector", "very-weak-vector"):
            service.ingest_document(text_document(f"test:{label}", label))

        response = service.search({"query": "controlled query", "limit": 100})
        self.assertEqual(2, response["candidate_count"])
        self.assertEqual(0, response["eligible_count"])
        self.assertEqual([], response["results"])
        self.assertAlmostEqual(0.4, response["cutoff_similarity"], places=6)

    def test_embedding_contract_remains_single_and_fingerprinted(self) -> None:
        self.service.ingest_document(document())
        with self.assertRaisesRegex(RuntimeError, "embedding contract"):
            SearchService(self.settings, embedder=ControlledEmbedder())

        class ChangedImplementation(ControlledEmbedder):
            fingerprint = "sha256:" + "3" * 64

        settings = Settings.explicit(Path(self.temporary.name) / "fingerprint-change")
        service = SearchService(settings, embedder=ControlledEmbedder())
        service.ingest_document(document())
        with self.assertRaisesRegex(RuntimeError, "embedding contract"):
            SearchService(settings, embedder=ChangedImplementation())

    def test_clear_index_allows_an_explicit_embedding_model_change(self) -> None:
        settings = Settings.explicit(Path(self.temporary.name) / "model-change")
        initial_service = SearchService(settings)
        initial_service.ingest_document(document())
        initial_service.clear_index()

        switched_service = SearchService(settings, embedder=ControlledEmbedder())
        switched_service.ingest_document(text_document("test:controlled", "controlled-query"))
        response = switched_service.search({"query": "controlled-query", "limit": 2})
        self.assertEqual("test:controlled", response["results"][0]["source_key"])
        self.assertEqual("controlled-test-v1", switched_service.stats()["embedding_model"])

    def test_vector_candidate_pool_keeps_best_chunk_per_document(self) -> None:
        settings = replace(
            Settings.explicit(Path(self.temporary.name) / "document-diversity"),
            chunk_chars=64,
            chunk_overlap_chars=0,
        )
        service = SearchService(settings, embedder=ControlledEmbedder())
        service.ingest_document(
            text_document("test:many-near-chunks", " ".join(["near-vector"] * 40))
        )
        service.ingest_document(text_document("test:one-close-chunk", "close-vector"))

        results = service.database.search(
            "controlled-query",
            limit=10,
            embedder=service.embedder,
            candidate_limit=2,
        )
        self.assertEqual(
            {"test:many-near-chunks", "test:one-close-chunk"},
            {item["source_key"] for item in results},
        )

    def test_lexical_candidate_pool_keeps_best_chunk_per_document(self) -> None:
        settings = replace(
            Settings.explicit(Path(self.temporary.name) / "lexical-diversity"),
            chunk_chars=64,
            chunk_overlap_chars=0,
        )
        service = SearchService(settings)
        service.ingest_document(
            text_document("test:many-literal-chunks", " ".join(["needle"] * 80))
        )
        service.ingest_document(
            text_document("test:one-literal-chunk", "needle once")
        )

        results = service.database.search(
            "needle",
            limit=10,
            embedder=service.embedder,
            candidate_limit=2,
        )
        self.assertEqual(
            {"test:many-literal-chunks", "test:one-literal-chunk"},
            {item["source_key"] for item in results},
        )

    def test_metadata_uses_normal_bounded_chunks(self) -> None:
        settings = replace(
            Settings.explicit(Path(self.temporary.name) / "metadata-chunks"),
            chunk_chars=64,
            chunk_overlap_chars=8,
        )
        service = SearchService(settings)
        service.ingest_document(
            document(metadata={"long": "metadata-token " * 100})
        )
        with service.database.session() as connection:
            rows = list(
                connection.execute(
                    "SELECT text FROM chunks WHERE part_kind = 'metadata'"
                )
            )
        self.assertGreater(len(rows), 1)
        self.assertTrue(all(len(row["text"]) <= 64 for row in rows))

    def test_metadata_search_projection_is_cropped_without_changing_stored_json(self) -> None:
        settings = replace(
            Settings.explicit(Path(self.temporary.name) / "metadata-projection-cap"),
            metadata_search_text_limit_chars=64,
            chunk_chars=64,
            chunk_overlap_chars=0,
        )
        service = SearchService(settings)
        metadata = {"long": "metadata-token " * 100}
        service.ingest_document(document(metadata=metadata, parts=[]))

        with service.database.session() as connection:
            row = connection.execute(
                "SELECT metadata_json FROM documents WHERE source_key = ?",
                ("test:document-1",),
            ).fetchone()
            indexed_characters = int(
                connection.execute(
                    "SELECT COALESCE(SUM(length(text)), 0) FROM chunks "
                    "WHERE part_kind = 'metadata'"
                ).fetchone()[0]
            )
        self.assertEqual(metadata, json.loads(row["metadata_json"]))
        self.assertLessEqual(indexed_characters, 64)

    def test_binary_extraction_has_an_aggregate_document_text_cap(self) -> None:
        settings = replace(
            Settings.explicit(Path(self.temporary.name) / "binary-aggregate-cap"),
            inline_part_limit_bytes=16,
            extracted_text_limit_chars=16,
            binary_extracted_text_limit_chars=5,
        )
        service = SearchService(settings)
        service.ingest_document(
            document(
                metadata={},
                parts=[
                    {
                        "key": "attachment:0",
                        "kind": "attachment",
                        "name": "first.txt",
                        "media_type": "text/plain",
                        "size": 4,
                        "content_base64": base64.b64encode(b"abcd").decode("ascii"),
                    },
                    {
                        "key": "attachment:1",
                        "kind": "attachment",
                        "name": "second.txt",
                        "media_type": "text/plain",
                        "size": 4,
                        "content_base64": base64.b64encode(b"efgh").decode("ascii"),
                    },
                ],
            )
        )

        with service.database.session() as connection:
            parts = list(
                connection.execute(
                    "SELECT part_key, extraction_status, extraction_error "
                    "FROM parts ORDER BY part_key"
                )
            )
            second_text = connection.execute(
                "SELECT text FROM chunks WHERE part_key = 'attachment:1'"
            ).fetchone()[0]
        self.assertEqual("extracted", parts[0]["extraction_status"])
        self.assertEqual("extracted_truncated", parts[1]["extraction_status"])
        self.assertIn("aggregate binary limit", parts[1]["extraction_error"])
        self.assertEqual("e", second_text)

    def test_many_small_derived_parts_are_cropped_instead_of_rejecting_document(self) -> None:
        settings = replace(
            Settings.explicit(Path(self.temporary.name) / "derived-chunk-budget"),
            document_text_limit_chars=1_000,
            metadata_search_text_limit_chars=1_000,
            binary_extracted_text_limit_chars=100,
            document_chunk_limit=4,
            chunk_chars=64,
            chunk_overlap_chars=0,
        )
        service = SearchService(settings)
        binary_parts = [
            {
                "key": f"attachment:{index}",
                "kind": "attachment",
                "name": f"part-{index}.txt",
                "media_type": "text/plain",
                "size": 1,
                "content_base64": base64.b64encode(b"x").decode("ascii"),
            }
            for index in range(10)
        ]
        service.ingest_document(
            document(
                metadata={"description": "derived metadata " * 20},
                parts=[
                    {
                        "key": "body",
                        "kind": "body",
                        "name": "body",
                        "media_type": "text/plain",
                        "size": 100,
                        "text": "body " * 20,
                    },
                    *binary_parts,
                ],
            )
        )

        with service.database.session() as connection:
            chunk_count = int(connection.execute("SELECT COUNT(*) FROM chunks").fetchone()[0])
            truncated = int(
                connection.execute(
                    "SELECT COUNT(*) FROM parts "
                    "WHERE extraction_status = 'extracted_truncated'"
                ).fetchone()[0]
            )
        self.assertEqual(4, chunk_count)
        self.assertGreater(truncated, 0)

    def test_document_text_chunk_and_query_budgets_are_enforced(self) -> None:
        settings = replace(
            Settings.explicit(Path(self.temporary.name) / "document-budgets"),
            document_text_limit_chars=100,
            document_chunk_limit=2,
            query_limit_chars=8,
            chunk_chars=64,
            chunk_overlap_chars=0,
        )
        service = SearchService(settings)
        with self.assertRaisesRegex(ValidationError, "searchable text exceeds 100"):
            service.ingest_document(
                text_document("test:text-budget", "x" * 101)
            )
        with self.assertRaisesRegex(ValidationError, "exceeds 2 chunks"):
            chunk_service = SearchService(
                replace(settings, document_text_limit_chars=10_000),
                embedder=service.embedder,
            )
            chunk_service.ingest_document(
                text_document("test:chunk-budget", "word " * 35)
            )
        with self.assertRaisesRegex(ValidationError, "query exceeds 8"):
            service.search({"query": "q" * 9, "limit": 1})

    def test_literal_chunk_supplies_snippet_when_vector_chunk_was_first(self) -> None:
        settings = replace(
            Settings.explicit(Path(self.temporary.name) / "literal-snippet"),
            chunk_chars=64,
            chunk_overlap_chars=0,
        )
        service = SearchService(settings, embedder=ControlledEmbedder())
        service.ingest_document(
            document(
                source_key="test:mixed-parts",
                metadata={},
                parts=[
                    {
                        "key": "semantic",
                        "kind": "body",
                        "name": "semantic",
                        "size": 0,
                        "text": "near-vector unrelated words",
                    },
                    {
                        "key": "literal",
                        "kind": "body",
                        "name": "literal",
                        "size": 0,
                        "text": "needle exact literal context",
                    },
                ],
            )
        )
        result = service.search({"query": "needle", "limit": 1})["results"][0]
        self.assertIn("needle", result["snippet"])
        self.assertEqual("literal", result["snippet_part"])

    def test_single_word_prefix_and_suffix_literal_hits_beat_dense_distractor(self) -> None:
        settings = Settings.explicit(Path(self.temporary.name) / "trigram-ranking")
        service = SearchService(settings, embedder=AdversarialShortQueryEmbedder())
        service.ingest_document(
            text_document("test:target-cybersport", "КИБЕРСПОРТ и длинная история переписки")
        )
        service.ingest_document(text_document("test:dense-distractor", "куку"))

        expected_kinds = {
            "киберспорт": "token",
            "кибер": "prefix",
            "спорт": "substring",
        }
        for query, expected_kind in expected_kinds.items():
            with self.subTest(query=query):
                response = service.search({"query": query, "limit": 25})
                self.assertTrue(response["lexical_gate"])
                self.assertEqual(
                    ["test:target-cybersport"],
                    [item["source_key"] for item in response["results"]],
                )
                self.assertEqual(expected_kind, response["results"][0]["lexical_match_kind"])

    def test_literal_hit_survives_full_vector_pool(self) -> None:
        settings = replace(
            Settings.explicit(Path(self.temporary.name) / "literal-vector-pool"),
            search_candidate_document_limit=2,
        )
        service = SearchService(settings, embedder=AdversarialShortQueryEmbedder())
        service.ingest_document(
            text_document("test:target-cybersport", "КИБЕРСПОРТ и длинная история переписки")
        )
        for index in range(3):
            service.ingest_document(text_document(f"test:dense-{index}", "куку"))

        response = service.search({"query": "кибер", "limit": 2})
        self.assertTrue(response["lexical_gate"])
        self.assertEqual(
            ["test:target-cybersport"],
            [item["source_key"] for item in response["results"]],
        )
        self.assertEqual("lexical_prefix", response["results"][0]["ranking_basis"])

    def test_schema_v3_is_rejected_without_modifying_it(self) -> None:
        settings = Settings.explicit(Path(self.temporary.name) / "schema-v3")
        service = SearchService(settings)
        with service.database.session() as connection:
            connection.execute("PRAGMA user_version = 3")

        with self.assertRaisesRegex(
            RuntimeError,
            "Unsupported database schema version 3; expected 4",
        ):
            service.database.initialize()
        with service.database.session() as connection:
            self.assertEqual(3, int(connection.execute("PRAGMA user_version").fetchone()[0]))

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
            data = json.dumps(payload, ensure_ascii=True).encode("utf-8")
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

    def test_health_advertises_protocol_and_v1_requires_token(self) -> None:
        status, payload = self._request("GET", "/health", token=False)
        self.assertEqual(200, status)
        self.assertEqual({"status": "ok", "protocol": 4}, payload)
        status, payload = self._request("GET", "/v1/stats", token=False)
        self.assertEqual(401, status)
        self.assertEqual("unauthorized", payload["error"])

    def test_messages_endpoint_is_removed(self) -> None:
        status, payload = self._request("POST", "/v1/messages", {"messages": []})
        self.assertEqual(404, status)
        self.assertEqual({"error": "not_found"}, payload)

    def test_http_document_ingest_and_generic_search_contract(self) -> None:
        locator = {"connector": "outlook_mapi", "store_id": "s", "entry_id": "e"}
        payload = document(locator=locator)
        status, outcome = self._request("POST", "/v1/documents", payload)
        self.assertEqual(200, status)
        self.assertEqual(
            {"source_key": "test:document-1", "status": "upserted"},
            outcome,
        )

        status, response = self._request(
            "POST", "/v1/search", {"query": "OriginalUniqueTerm", "limit": 5}
        )
        self.assertEqual(200, status)
        result = response["results"][0]
        self.assertEqual("test:document-1", result["source_key"])
        self.assertEqual(payload["metadata"], result["metadata"])
        self.assertEqual(locator, result["locator"])
        self.assertEqual("body", result["snippet_part"])
        self.assertIn("body", result["matched_parts"])
        self.assertEqual(1, result["rank"])
        self.assertIn("vector_distance", result)

    def test_invalid_inline_content_returns_bad_request(self) -> None:
        payload = document(
            parts=[
                {
                    "key": "attachment",
                    "kind": "attachment",
                    "size": 1,
                    "content_base64": "?",
                }
            ]
        )
        status, response = self._request("POST", "/v1/documents", payload)
        self.assertEqual(400, status)
        self.assertEqual("invalid_request", response["error"])

    def test_unpaired_surrogate_returns_bad_request(self) -> None:
        status, response = self._request(
            "POST",
            "/v1/documents",
            document(title="bad\ud800title"),
        )
        self.assertEqual(400, status)
        self.assertEqual("invalid_request", response["error"])
        self.assertIn("unpaired UTF-16 surrogate", response["message"])

    def test_opaque_integer_outside_host_double_range_returns_bad_request(self) -> None:
        status, response = self._request(
            "POST",
            "/v1/documents",
            document(metadata={"number": 10**1000}),
        )
        self.assertEqual(400, status)
        self.assertEqual("invalid_request", response["error"])
        self.assertIn("finite IEEE-754 range", response["message"])

    def test_clear_index_uses_generic_counters_and_service_remains_usable(self) -> None:
        self.service.ingest_document(document())
        before = self.service.stats()
        original_token = self.service.token

        status, payload = self._request("DELETE", "/v1/index")
        self.assertEqual(200, status)
        self.assertEqual(before["documents"], payload["deleted_documents"])
        self.assertEqual(before["parts"], payload["deleted_parts"])
        self.assertEqual(before["chunks"], payload["deleted_chunks"])
        after = self.service.stats()
        self.assertEqual(0, after["documents"])
        self.assertEqual(0, after["parts"])
        self.assertEqual(0, after["chunks"])
        self.assertIsNone(after["embedding_model"])
        self.assertEqual(original_token, self.service.token)

        status, payload = self._request("DELETE", "/v1/index")
        self.assertEqual(200, status)
        self.assertEqual(
            {"deleted_documents": 0, "deleted_parts": 0, "deleted_chunks": 0},
            payload,
        )
        status, outcome = self._request("POST", "/v1/documents", document())
        self.assertEqual(200, status)
        self.assertEqual("upserted", outcome["status"])


if __name__ == "__main__":
    unittest.main()
