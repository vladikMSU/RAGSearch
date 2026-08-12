from __future__ import annotations

import hashlib
import heapq
import json
import re
import sqlite3
from collections.abc import Iterable, Sequence
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .chunking import normalize_text
from .embeddings import Embedder, blob_to_vector, cosine_for_normalized, vector_to_blob


SCHEMA_VERSION = 4
_SEARCH_TOKEN = re.compile(r"[\w@.+-]+", re.UNICODE)


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _sha256(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


class Database:
    def __init__(self, path: Path) -> None:
        self.path = Path(path)

    def connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.path, timeout=30.0)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA busy_timeout = 30000")
        return connection

    @contextmanager
    def session(self) -> Iterable[sqlite3.Connection]:
        connection = self.connect()
        try:
            with connection:
                yield connection
        finally:
            connection.close()

    def initialize(self) -> None:
        fresh_database = not self.path.exists()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.session() as connection:
            schema_version = int(
                connection.execute("PRAGMA user_version").fetchone()[0]
            )
            if schema_version != SCHEMA_VERSION and not (
                fresh_database and schema_version == 0
            ):
                raise RuntimeError(
                    "Unsupported database schema version "
                    f"{schema_version}; expected {SCHEMA_VERSION}. "
                    "Delete the database file and rebuild the local index."
                )
            connection.execute("PRAGMA journal_mode = WAL")
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS documents (
                    id INTEGER PRIMARY KEY,
                    source_key TEXT NOT NULL UNIQUE,
                    kind TEXT NOT NULL,
                    title TEXT NOT NULL DEFAULT '',
                    metadata_json TEXT NOT NULL,
                    locator_json TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE INDEX IF NOT EXISTS idx_documents_kind
                    ON documents (kind);

                CREATE TABLE IF NOT EXISTS parts (
                    id INTEGER PRIMARY KEY,
                    document_id INTEGER NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
                    part_key TEXT NOT NULL,
                    kind TEXT NOT NULL,
                    name TEXT NOT NULL DEFAULT '',
                    media_type TEXT NOT NULL DEFAULT '',
                    declared_size INTEGER NOT NULL DEFAULT 0,
                    truncated INTEGER NOT NULL DEFAULT 0 CHECK (truncated IN (0, 1)),
                    extraction_status TEXT NOT NULL,
                    extraction_error TEXT,
                    text_hash TEXT NOT NULL DEFAULT '',
                    UNIQUE (document_id, part_key)
                );

                CREATE INDEX IF NOT EXISTS idx_parts_document
                    ON parts (document_id);

                CREATE TABLE IF NOT EXISTS chunks (
                    id INTEGER PRIMARY KEY,
                    document_id INTEGER NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
                    part_id INTEGER REFERENCES parts(id) ON DELETE CASCADE,
                    part_key TEXT NOT NULL,
                    part_kind TEXT NOT NULL,
                    part_label TEXT NOT NULL,
                    ordinal INTEGER NOT NULL,
                    text TEXT NOT NULL,
                    text_hash TEXT NOT NULL,
                    embedding BLOB NOT NULL,
                    embedding_model TEXT NOT NULL,
                    embedding_dim INTEGER NOT NULL,
                    embedding_fingerprint TEXT NOT NULL,
                    UNIQUE (document_id, part_key, ordinal)
                );

                CREATE INDEX IF NOT EXISTS idx_chunks_document
                    ON chunks (document_id);
                CREATE INDEX IF NOT EXISTS idx_chunks_embedding
                    ON chunks (embedding_model, embedding_dim, embedding_fingerprint);

                CREATE VIRTUAL TABLE IF NOT EXISTS chunks_fts USING fts5(
                    text,
                    content='chunks',
                    content_rowid='id',
                    tokenize='unicode61 remove_diacritics 2'
                );

                CREATE VIRTUAL TABLE IF NOT EXISTS chunks_trigram USING fts5(
                    text,
                    content='chunks',
                    content_rowid='id',
                    tokenize='trigram'
                );

                CREATE TRIGGER IF NOT EXISTS chunks_ai AFTER INSERT ON chunks BEGIN
                    INSERT INTO chunks_fts(rowid, text) VALUES (new.id, new.text);
                END;
                CREATE TRIGGER IF NOT EXISTS chunks_ad AFTER DELETE ON chunks BEGIN
                    INSERT INTO chunks_fts(chunks_fts, rowid, text)
                    VALUES ('delete', old.id, old.text);
                END;
                CREATE TRIGGER IF NOT EXISTS chunks_au AFTER UPDATE ON chunks BEGIN
                    INSERT INTO chunks_fts(chunks_fts, rowid, text)
                    VALUES ('delete', old.id, old.text);
                    INSERT INTO chunks_fts(rowid, text) VALUES (new.id, new.text);
                END;

                CREATE TRIGGER IF NOT EXISTS chunks_trigram_ai AFTER INSERT ON chunks BEGIN
                    INSERT INTO chunks_trigram(rowid, text) VALUES (new.id, new.text);
                END;
                CREATE TRIGGER IF NOT EXISTS chunks_trigram_ad AFTER DELETE ON chunks BEGIN
                    INSERT INTO chunks_trigram(chunks_trigram, rowid, text)
                    VALUES ('delete', old.id, old.text);
                END;
                CREATE TRIGGER IF NOT EXISTS chunks_trigram_au AFTER UPDATE ON chunks BEGIN
                    INSERT INTO chunks_trigram(chunks_trigram, rowid, text)
                    VALUES ('delete', old.id, old.text);
                    INSERT INTO chunks_trigram(rowid, text) VALUES (new.id, new.text);
                END;

                CREATE TABLE IF NOT EXISTS service_meta (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
                """
            )
            connection.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")

    def ping(self) -> bool:
        with self.session() as connection:
            return connection.execute("SELECT 1").fetchone()[0] == 1

    @staticmethod
    def _validate_embedding_contract(
        connection: sqlite3.Connection,
        embedder: Embedder,
    ) -> bool:
        meta = {
            row["key"]: row["value"]
            for row in connection.execute(
                "SELECT key, value FROM service_meta "
                "WHERE key IN ('embedding_model', 'embedding_dim', 'embedding_fingerprint')"
            )
        }
        indexed_models = list(
            connection.execute(
                "SELECT DISTINCT embedding_model, embedding_dim, embedding_fingerprint "
                "FROM chunks LIMIT 2"
            )
        )
        if not meta and not indexed_models:
            return False

        expected_model = embedder.name
        expected_dim = str(embedder.dimensions)
        expected_fingerprint = embedder.fingerprint
        valid_meta = (
            meta.get("embedding_model") == expected_model
            and meta.get("embedding_dim") == expected_dim
            and meta.get("embedding_fingerprint") == expected_fingerprint
        )
        valid_chunks = not indexed_models or (
            len(indexed_models) == 1
            and indexed_models[0]["embedding_model"] == expected_model
            and str(indexed_models[0]["embedding_dim"]) == expected_dim
            and indexed_models[0]["embedding_fingerprint"] == expected_fingerprint
        )
        if not valid_meta or not valid_chunks:
            actual_models = ", ".join(
                f"{row['embedding_model']}:{row['embedding_dim']}:{row['embedding_fingerprint']}"
                for row in indexed_models
            ) or "no chunks"
            raise RuntimeError(
                "The index embedding contract does not match the configured embedder "
                f"{expected_model}:{expected_dim}:{expected_fingerprint} "
                f"(stored: {actual_models}). "
                "Clear the index and re-ingest it with one embedding model."
            )
        return True

    def validate_embedding_model(self, embedder: Embedder) -> None:
        with self.session() as connection:
            self._validate_embedding_contract(connection, embedder)

    def clear_index(self) -> dict[str, int]:
        """Delete indexed documents atomically while keeping the database itself."""
        with self.session() as connection:
            connection.execute("BEGIN IMMEDIATE")
            counts = {
                table: int(connection.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0])
                for table in ("documents", "parts", "chunks")
            }
            connection.execute("DELETE FROM chunks")
            connection.execute("DELETE FROM parts")
            connection.execute("DELETE FROM documents")
            connection.execute("INSERT INTO chunks_fts(chunks_fts) VALUES ('rebuild')")
            connection.execute(
                "INSERT INTO chunks_trigram(chunks_trigram) VALUES ('rebuild')"
            )
            connection.execute("DELETE FROM service_meta")

        return {
            "deleted_documents": counts["documents"],
            "deleted_parts": counts["parts"],
            "deleted_chunks": counts["chunks"],
        }

    def upsert_document(
        self,
        document: dict[str, Any],
        parts: Sequence[dict[str, Any]],
        chunks: Sequence[dict[str, Any]],
        *,
        embedder: Embedder,
    ) -> None:
        now = _utc_now()
        with self.session() as connection:
            connection.execute("BEGIN IMMEDIATE")
            has_embedding_contract = self._validate_embedding_contract(
                connection,
                embedder,
            )
            if not has_embedding_contract:
                connection.execute(
                    "INSERT INTO service_meta(key, value) VALUES ('embedding_model', ?)",
                    (embedder.name,),
                )
                connection.execute(
                    "INSERT INTO service_meta(key, value) VALUES ('embedding_dim', ?)",
                    (str(embedder.dimensions),),
                )
                connection.execute(
                    "INSERT INTO service_meta(key, value) "
                    "VALUES ('embedding_fingerprint', ?)",
                    (embedder.fingerprint,),
                )

            connection.execute(
                """
                INSERT INTO documents (
                    source_key, kind, title, metadata_json, locator_json,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(source_key) DO UPDATE SET
                    kind=excluded.kind,
                    title=excluded.title,
                    metadata_json=excluded.metadata_json,
                    locator_json=excluded.locator_json,
                    updated_at=excluded.updated_at
                """,
                (
                    document["source_key"],
                    document["kind"],
                    document["title"],
                    document["metadata_json"],
                    document["locator_json"],
                    now,
                    now,
                ),
            )
            document_id = int(
                connection.execute(
                    "SELECT id FROM documents WHERE source_key = ?",
                    (document["source_key"],),
                ).fetchone()[0]
            )

            # Each request is a complete document snapshot.  Replacing children
            # makes retries idempotent and removes parts missing from a later snapshot.
            connection.execute("DELETE FROM chunks WHERE document_id = ?", (document_id,))
            connection.execute("DELETE FROM parts WHERE document_id = ?", (document_id,))

            part_ids: dict[str, int] = {}
            for part in parts:
                cursor = connection.execute(
                    """
                    INSERT INTO parts (
                        document_id, part_key, kind, name, media_type, declared_size,
                        truncated, extraction_status, extraction_error, text_hash
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        document_id,
                        part["part_key"],
                        part["kind"],
                        part["name"],
                        part["media_type"],
                        part["size"],
                        int(part["truncated"]),
                        part["extraction_status"],
                        part["extraction_error"],
                        part["text_hash"],
                    ),
                )
                part_ids[str(part["part_key"])] = int(cursor.lastrowid)

            for chunk in chunks:
                part_key = str(chunk["part_key"])
                part_id = part_ids.get(part_key) if part_key else None
                connection.execute(
                    """
                    INSERT INTO chunks (
                        document_id, part_id, part_key, part_kind, part_label,
                        ordinal, text, text_hash, embedding, embedding_model,
                        embedding_dim, embedding_fingerprint
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        document_id,
                        part_id,
                        part_key,
                        chunk["part_kind"],
                        chunk["part_label"],
                        chunk["ordinal"],
                        chunk["text"],
                        _sha256(chunk["text"]),
                        vector_to_blob(chunk["embedding"]),
                        embedder.name,
                        embedder.dimensions,
                        embedder.fingerprint,
                    ),
                )

    @staticmethod
    def _fts_expression(query: str) -> str:
        tokens: list[str] = []
        seen: set[str] = set()
        for match in _SEARCH_TOKEN.finditer(query):
            token = match.group(0).casefold()
            if token not in seen:
                seen.add(token)
                tokens.append(token)
            if len(tokens) >= 32:
                break
        return " OR ".join('"' + token.replace('"', '""') + '"' for token in tokens)

    @staticmethod
    def _trigram_expression(query: str) -> str:
        tokens: list[str] = []
        seen: set[str] = set()
        for match in _SEARCH_TOKEN.finditer(query):
            token = match.group(0).casefold()
            if len(token) >= 3 and token not in seen:
                seen.add(token)
                tokens.append(token)
            if len(tokens) >= 32:
                break
        return " AND ".join('"' + token.replace('"', '""') + '"' for token in tokens)

    @staticmethod
    def _prefix_expression(query: str) -> str:
        tokens: list[str] = []
        seen: set[str] = set()
        for match in _SEARCH_TOKEN.finditer(query):
            token = match.group(0).casefold()
            if len(token) >= 3 and token not in seen:
                seen.add(token)
                tokens.append(token)
            if len(tokens) >= 32:
                break
        return " AND ".join(
            '"' + token.replace('"', '""') + '"*' for token in tokens
        )

    @staticmethod
    def _lexical_kind_priority(kind: str) -> int:
        return {"token": 3, "prefix": 2, "substring": 1}.get(kind, 0)

    @staticmethod
    def _row_payload(row: sqlite3.Row) -> dict[str, Any]:
        return {
            "chunk_id": row["chunk_id"],
            "document_id": row["document_id"],
            "text": row["text"],
            "part_kind": row["part_kind"],
            "part_label": row["part_label"],
        }

    @staticmethod
    def _snippet(text: str, query: str, max_chars: int = 280) -> str:
        flattened = " ".join(normalize_text(text).split())
        if len(flattened) <= max_chars:
            return flattened
        folded = flattened.casefold()
        probes = [query.casefold()] + [
            match.group(0).casefold() for match in _SEARCH_TOKEN.finditer(query)
        ]
        position = next((folded.find(probe) for probe in probes if folded.find(probe) >= 0), 0)
        start = max(0, position - max_chars // 3)
        end = min(len(flattened), start + max_chars)
        prefix = "…" if start else ""
        suffix = "…" if end < len(flattened) else ""
        return prefix + flattened[start:end].strip() + suffix

    def search(
        self,
        query: str,
        *,
        limit: int,
        embedder: Embedder,
        candidate_limit: int,
    ) -> list[dict[str, Any]]:
        query_vector = embedder.embed_many([query])[0]
        candidates: dict[int, dict[str, Any]] = {}
        select_fields = """
            c.id AS chunk_id, c.document_id, c.text, c.part_kind, c.part_label
        """

        with self.session() as connection:
            # Keep chunk selection and the later document projection on one WAL
            # read snapshot while ingestion/reset may write concurrently.
            connection.execute("BEGIN")
            fts_expression = self._fts_expression(query)
            prefix_expression = self._prefix_expression(query)
            trigram_expression = self._trigram_expression(query)
            lexical_rows: dict[int, tuple[sqlite3.Row, float, str]] = {}
            for table_name, expression, match_kind in (
                ("chunks_fts", fts_expression, "token"),
                ("chunks_fts", prefix_expression, "prefix"),
                ("chunks_trigram", trigram_expression, "substring"),
            ):
                if not expression:
                    continue
                lexical_sql = f"""
                    SELECT {select_fields}, c.embedding, c.embedding_model, c.embedding_dim,
                           c.embedding_fingerprint,
                           bm25({table_name}) AS lexical_rank
                    FROM {table_name}
                    JOIN chunks c ON c.id = {table_name}.rowid
                    WHERE {table_name} MATCH ?
                    ORDER BY lexical_rank
                    LIMIT ?
                """
                # FTS5's bm25() cannot be used through a window/CTE context.
                # Over-fetch chunks, then retain the best one per document before
                # applying the document candidate limit in Python.
                chunk_limit = min(candidate_limit * 64, 100_000)
                rows: list[sqlite3.Row] = []
                seen_documents: set[int] = set()
                for row in connection.execute(lexical_sql, [expression, chunk_limit]):
                    document_id = int(row["document_id"])
                    if document_id in seen_documents:
                        continue
                    seen_documents.add(document_id)
                    rows.append(row)
                    if len(rows) >= candidate_limit:
                        break
                maximum_score = max(
                    (max(0.0, -float(row["lexical_rank"])) for row in rows),
                    default=1.0,
                ) or 1.0
                for row in rows:
                    score = max(0.0, -float(row["lexical_rank"])) / maximum_score
                    chunk_id = int(row["chunk_id"])
                    current = lexical_rows.get(chunk_id)
                    if current is None or score > current[1] or (
                        score == current[1]
                        and self._lexical_kind_priority(match_kind)
                        > self._lexical_kind_priority(current[2])
                    ):
                        lexical_rows[chunk_id] = (row, score, match_kind)

            for row, lexical_score, match_kind in lexical_rows.values():
                payload = self._row_payload(row)
                payload["lexical"] = lexical_score
                payload["lexical_match_kind"] = match_kind
                payload["vector"] = 0.0
                payload["vector_available"] = False
                try:
                    stored = blob_to_vector(row["embedding"], int(row["embedding_dim"]))
                    if (
                        row["embedding_model"] == embedder.name
                        and row["embedding_fingerprint"] == embedder.fingerprint
                        and len(stored) == len(query_vector)
                    ):
                        payload["vector"] = min(
                            1.0,
                            max(0.0, cosine_for_normalized(query_vector, stored)),
                        )
                        payload["vector_available"] = True
                except (TypeError, ValueError):
                    pass
                candidates[int(row["chunk_id"])] = payload

            vector_sql = f"""
                SELECT {select_fields}, c.embedding
                FROM chunks c
                WHERE c.embedding_model = ? AND c.embedding_dim = ?
                  AND c.embedding_fingerprint = ?
            """
            vector_parameters = [
                embedder.name,
                embedder.dimensions,
                embedder.fingerprint,
            ]
            best_vector_chunk_by_document: dict[
                int, tuple[float, int, sqlite3.Row]
            ] = {}
            for row in connection.execute(vector_sql, vector_parameters):
                try:
                    stored = blob_to_vector(row["embedding"], embedder.dimensions)
                except (TypeError, ValueError):
                    continue
                similarity = min(
                    1.0,
                    max(0.0, cosine_for_normalized(query_vector, stored)),
                )
                item = (similarity, int(row["chunk_id"]), row)
                document_id = int(row["document_id"])
                current = best_vector_chunk_by_document.get(document_id)
                if current is None or item[:2] > current[:2]:
                    best_vector_chunk_by_document[document_id] = item

            vector_candidates = heapq.nlargest(
                candidate_limit,
                best_vector_chunk_by_document.values(),
                key=lambda item: item[:2],
            )
            for similarity, chunk_id, row in vector_candidates:
                existing = candidates.get(chunk_id)
                if existing is None:
                    existing = self._row_payload(row)
                    existing["lexical"] = 0.0
                    existing["lexical_match_kind"] = ""
                    candidates[chunk_id] = existing
                existing["vector"] = max(float(existing.get("vector", 0.0)), similarity)
                existing["vector_available"] = True

            # Fetch the opaque document projection once per candidate document;
            # never duplicate large metadata/locator JSON across chunk rows.
            candidate_document_ids = sorted(
                {int(item["document_id"]) for item in candidates.values()}
            )
            documents: dict[int, sqlite3.Row] = {}
            if candidate_document_ids:
                placeholders = ",".join("?" for _ in candidate_document_ids)
                for row in connection.execute(
                    "SELECT id, source_key, kind, title, metadata_json, locator_json "
                    f"FROM documents WHERE id IN ({placeholders})",
                    candidate_document_ids,
                ):
                    documents[int(row["id"])] = row
            for candidate in candidates.values():
                document = documents[int(candidate["document_id"])]
                candidate["source_key"] = document["source_key"]
                candidate["kind"] = document["kind"]
                candidate["title"] = document["title"]
                candidate["metadata_search"] = document["metadata_json"]
                candidate["metadata_json"] = document["metadata_json"]
                candidate["locator_json"] = document["locator_json"]

        folded_query = query.casefold()
        ranked_chunks: list[tuple[float, dict[str, Any]]] = []
        for candidate in candidates.values():
            score = 0.62 * candidate["vector"] + 0.38 * candidate["lexical"]
            searchable_metadata = (
                str(candidate["title"]) + " " + str(candidate["metadata_search"])
            ).casefold()
            if folded_query in searchable_metadata:
                score += 0.12
            ranked_chunks.append((min(1.0, score), candidate))
        ranked_chunks.sort(
            key=lambda item: (-item[0], item[1]["document_id"], item[1]["chunk_id"])
        )

        by_document: dict[int, dict[str, Any]] = {}
        for score, candidate in ranked_chunks:
            document_id = int(candidate["document_id"])
            part = candidate["part_label"]
            result = by_document.get(document_id)
            if result is None:
                result = {
                    "source_key": candidate["source_key"],
                    "kind": candidate["kind"],
                    "title": candidate["title"],
                    "metadata": json.loads(candidate["metadata_json"]),
                    "locator": json.loads(candidate["locator_json"]),
                    "vector_similarity": float(candidate["vector"]),
                    "vector_available": bool(candidate.get("vector_available", False)),
                    "lexical_score": float(candidate["lexical"]),
                    "lexical_match_kind": str(
                        candidate.get("lexical_match_kind", "")
                    ),
                    "hybrid_score": float(score),
                    "snippet": self._snippet(candidate["text"], query),
                    "snippet_part": part,
                    "matched_parts": [part],
                }
                by_document[document_id] = result
            else:
                result["vector_similarity"] = max(
                    float(result["vector_similarity"]),
                    float(candidate["vector"]),
                )
                result["vector_available"] = bool(result["vector_available"]) or bool(
                    candidate.get("vector_available", False)
                )
                candidate_lexical = float(candidate["lexical"])
                current_lexical = float(result["lexical_score"])
                candidate_kind = str(candidate.get("lexical_match_kind", ""))
                if candidate_lexical > current_lexical or (
                    candidate_lexical == current_lexical
                    and self._lexical_kind_priority(candidate_kind)
                    > self._lexical_kind_priority(
                        str(result["lexical_match_kind"])
                    )
                ):
                    result["lexical_score"] = candidate_lexical
                    result["lexical_match_kind"] = candidate_kind
                    if candidate_lexical > 0.0:
                        result["snippet"] = self._snippet(candidate["text"], query)
                        result["snippet_part"] = part
                result["hybrid_score"] = max(
                    float(result["hybrid_score"]),
                    float(score),
                )
                if part not in result["matched_parts"] and len(result["matched_parts"]) < 8:
                    result["matched_parts"].append(part)

        results = list(by_document.values())
        for result in results:
            similarity = min(1.0, max(0.0, float(result["vector_similarity"])))
            hybrid_score = min(1.0, max(0.0, float(result["hybrid_score"])))
            result["vector_similarity"] = round(similarity, 6)
            result["vector_distance"] = round(1.0 - similarity, 6)
            result["lexical_score"] = round(
                min(1.0, max(0.0, float(result["lexical_score"]))),
                6,
            )
            result["hybrid_score"] = round(hybrid_score, 6)
        results.sort(
            key=lambda result: (
                0 if bool(result["vector_available"]) else 1,
                float(result["vector_distance"]),
                -float(result["hybrid_score"]),
                str(result["source_key"]),
            )
        )
        vector_results = [result for result in results if bool(result["vector_available"])]
        lexical_results = [
            result
            for result in results
            if float(result["lexical_score"]) > 0.0
        ]
        lexical_results.sort(
            key=lambda result: (
                -self._lexical_kind_priority(str(result["lexical_match_kind"])),
                -float(result["lexical_score"]),
                0 if bool(result["vector_available"]) else 1,
                float(result["vector_distance"]),
                -float(result["hybrid_score"]),
                str(result["source_key"]),
            )
        )

        # Keep independently bounded vector and lexical pools. A literal hit can
        # have weak cosine similarity and must reach the service-level lexical gate.
        merged: list[dict[str, Any]] = []
        seen_source_keys: set[str] = set()
        for result in vector_results[:limit] + lexical_results[:limit]:
            identity = str(result["source_key"])
            if identity in seen_source_keys:
                continue
            seen_source_keys.add(identity)
            merged.append(result)
        return merged

    def stats(self) -> dict[str, Any]:
        with self.session() as connection:
            counts = {
                table: int(connection.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0])
                for table in ("documents", "parts", "chunks")
            }
            extraction = {
                row["extraction_status"]: int(row["amount"])
                for row in connection.execute(
                    """
                    SELECT extraction_status, COUNT(*) AS amount
                    FROM parts GROUP BY extraction_status
                    """
                )
            }
            meta = {
                row["key"]: row["value"]
                for row in connection.execute("SELECT key, value FROM service_meta")
            }
        try:
            database_bytes = self.path.stat().st_size
        except OSError:
            database_bytes = 0
        return {
            **counts,
            "part_extraction": extraction,
            "database_bytes": database_bytes,
            "schema_version": SCHEMA_VERSION,
            "embedding_model": meta.get("embedding_model"),
            "embedding_dim": int(meta["embedding_dim"]) if "embedding_dim" in meta else None,
            "embedding_fingerprint": meta.get("embedding_fingerprint"),
        }
