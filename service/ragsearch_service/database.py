from __future__ import annotations

import hashlib
import heapq
import re
import sqlite3
from contextlib import contextmanager
from collections.abc import Iterable, Sequence
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .chunking import normalize_text
from .embeddings import Embedder, blob_to_vector, cosine_for_normalized, vector_to_blob
from .errors import ValidationError
from .timestamps import canonical_utc_timestamp


SCHEMA_VERSION = 3
_SEARCH_TOKEN = re.compile(r"[\w@.+-]+", re.UNICODE)


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _sha256(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _escape_like(value: str) -> str:
    return value.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


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
                CREATE TABLE IF NOT EXISTS messages (
                    id INTEGER PRIMARY KEY,
                    entry_id TEXT NOT NULL,
                    store_id TEXT NOT NULL,
                    folder_entry_id TEXT NOT NULL,
                    folder_path TEXT NOT NULL DEFAULT '',
                    store_name TEXT NOT NULL DEFAULT '',
                    subject TEXT NOT NULL DEFAULT '',
                    sender_name TEXT NOT NULL DEFAULT '',
                    sender_email TEXT NOT NULL DEFAULT '',
                    to_recipients TEXT NOT NULL DEFAULT '',
                    cc_recipients TEXT NOT NULL DEFAULT '',
                    sent_at TEXT,
                    received_at TEXT,
                    modified_at TEXT,
                    internet_message_id TEXT NOT NULL DEFAULT '',
                    conversation_id TEXT NOT NULL DEFAULT '',
                    body_hash TEXT NOT NULL DEFAULT '',
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE (store_id, entry_id)
                );

                CREATE INDEX IF NOT EXISTS idx_messages_folder
                    ON messages (store_id, folder_entry_id);
                CREATE INDEX IF NOT EXISTS idx_messages_received
                    ON messages (received_at);
                CREATE INDEX IF NOT EXISTS idx_messages_sender
                    ON messages (sender_email);

                CREATE TABLE IF NOT EXISTS attachments (
                    id INTEGER PRIMARY KEY,
                    message_id INTEGER NOT NULL REFERENCES messages(id) ON DELETE CASCADE,
                    attachment_index INTEGER NOT NULL,
                    name TEXT NOT NULL DEFAULT '',
                    declared_size INTEGER NOT NULL DEFAULT 0,
                    content_type TEXT NOT NULL DEFAULT '',
                    extraction_status TEXT NOT NULL,
                    extraction_error TEXT,
                    text_hash TEXT NOT NULL DEFAULT '',
                    UNIQUE (message_id, attachment_index)
                );

                CREATE INDEX IF NOT EXISTS idx_attachments_message
                    ON attachments (message_id);

                CREATE TABLE IF NOT EXISTS chunks (
                    id INTEGER PRIMARY KEY,
                    message_id INTEGER NOT NULL REFERENCES messages(id) ON DELETE CASCADE,
                    attachment_id INTEGER REFERENCES attachments(id) ON DELETE CASCADE,
                    source_key TEXT NOT NULL,
                    source_kind TEXT NOT NULL CHECK (source_kind IN ('metadata', 'body', 'attachment')),
                    source_label TEXT NOT NULL,
                    ordinal INTEGER NOT NULL,
                    text TEXT NOT NULL,
                    text_hash TEXT NOT NULL,
                    embedding BLOB NOT NULL,
                    embedding_model TEXT NOT NULL,
                    embedding_dim INTEGER NOT NULL,
                    embedding_fingerprint TEXT NOT NULL,
                    UNIQUE (message_id, source_key, ordinal)
                );

                CREATE INDEX IF NOT EXISTS idx_chunks_message
                    ON chunks (message_id);
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
        """Delete indexed Outlook data atomically while keeping the database itself."""
        with self.session() as connection:
            connection.execute("BEGIN IMMEDIATE")
            counts = {
                table: int(connection.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0])
                for table in ("messages", "attachments", "chunks")
            }

            # Delete children explicitly so the FTS delete trigger runs for every
            # chunk. Rebuild then guarantees that the external-content FTS index
            # exactly reflects the now-empty chunks table.
            connection.execute("DELETE FROM chunks")
            connection.execute("DELETE FROM attachments")
            connection.execute("DELETE FROM messages")
            connection.execute("INSERT INTO chunks_fts(chunks_fts) VALUES ('rebuild')")
            connection.execute(
                "INSERT INTO chunks_trigram(chunks_trigram) VALUES ('rebuild')"
            )
            connection.execute("DELETE FROM service_meta")

        return {
            "deleted_messages": counts["messages"],
            "deleted_attachments": counts["attachments"],
            "deleted_chunks": counts["chunks"],
        }

    def upsert_message(
        self,
        message: dict[str, Any],
        attachments: Sequence[dict[str, Any]],
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
                INSERT INTO messages (
                    entry_id, store_id, folder_entry_id, folder_path, store_name,
                    subject, sender_name, sender_email, to_recipients, cc_recipients,
                    sent_at, received_at, modified_at, internet_message_id,
                    conversation_id, body_hash, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(store_id, entry_id) DO UPDATE SET
                    folder_entry_id=excluded.folder_entry_id,
                    folder_path=excluded.folder_path,
                    store_name=excluded.store_name,
                    subject=excluded.subject,
                    sender_name=excluded.sender_name,
                    sender_email=excluded.sender_email,
                    to_recipients=excluded.to_recipients,
                    cc_recipients=excluded.cc_recipients,
                    sent_at=excluded.sent_at,
                    received_at=excluded.received_at,
                    modified_at=excluded.modified_at,
                    internet_message_id=excluded.internet_message_id,
                    conversation_id=excluded.conversation_id,
                    body_hash=excluded.body_hash,
                    updated_at=excluded.updated_at
                """,
                (
                    message["entry_id"],
                    message["store_id"],
                    message["folder_entry_id"],
                    message["folder_path"],
                    message["store_name"],
                    message["subject"],
                    message["sender_name"],
                    message["sender_email"],
                    message["to"],
                    message["cc"],
                    message["sent_at"],
                    message["received_at"],
                    message["modified_at"],
                    message["internet_message_id"],
                    message["conversation_id"],
                    message["body_hash"],
                    now,
                    now,
                ),
            )
            message_id = connection.execute(
                "SELECT id FROM messages WHERE store_id = ? AND entry_id = ?",
                (message["store_id"], message["entry_id"]),
            ).fetchone()[0]

            # Every API item is a complete message snapshot. Replacing its children
            # makes retries idempotent and removes attachments deleted in Outlook.
            connection.execute("DELETE FROM chunks WHERE message_id = ?", (message_id,))
            connection.execute("DELETE FROM attachments WHERE message_id = ?", (message_id,))

            attachment_ids: dict[int, int] = {}
            for attachment in attachments:
                cursor = connection.execute(
                    """
                    INSERT INTO attachments (
                        message_id, attachment_index, name, declared_size, content_type,
                        extraction_status, extraction_error, text_hash
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        message_id,
                        attachment["attachment_index"],
                        attachment["name"],
                        attachment["size"],
                        attachment["content_type"],
                        attachment["extraction_status"],
                        attachment["extraction_error"],
                        attachment["text_hash"],
                    ),
                )
                attachment_ids[attachment["attachment_index"]] = int(cursor.lastrowid)

            for chunk in chunks:
                attachment_index = chunk.get("attachment_index")
                attachment_id = (
                    attachment_ids[int(attachment_index)] if attachment_index is not None else None
                )
                connection.execute(
                    """
                    INSERT INTO chunks (
                        message_id, attachment_id, source_key, source_kind, source_label,
                        ordinal, text, text_hash, embedding, embedding_model,
                        embedding_dim, embedding_fingerprint
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        message_id,
                        attachment_id,
                        chunk["source_key"],
                        chunk["source_kind"],
                        chunk["source_label"],
                        chunk["ordinal"],
                        chunk["text"],
                        _sha256(chunk["text"]),
                        vector_to_blob(chunk["embedding"]),
                        embedder.name,
                        embedder.dimensions,
                        embedder.fingerprint,
                    ),
                )

    def _filters(self, filters: object) -> tuple[str, list[object]]:
        if filters is None:
            return "", []
        if not isinstance(filters, dict):
            raise ValidationError("filters must be an object")
        allowed = {
            "store_id",
            "store_ids",
            "folder_entry_id",
            "folder_path",
            "folder_path_prefix",
            "sender_email",
            "received_from",
            "received_to",
            "has_attachments",
        }
        unknown = set(filters) - allowed
        if unknown:
            raise ValidationError(f"Unknown search filters: {', '.join(sorted(unknown))}")

        clauses: list[str] = []
        parameters: list[object] = []

        for name in ("store_id", "folder_entry_id", "folder_path"):
            value = filters.get(name)
            if value is not None:
                if not isinstance(value, str):
                    raise ValidationError(f"filters.{name} must be a string")
                clauses.append(f"m.{name} = ?")
                parameters.append(value)

        store_ids = filters.get("store_ids")
        if store_ids is not None:
            if not isinstance(store_ids, list) or not store_ids or not all(
                isinstance(item, str) and item for item in store_ids
            ):
                raise ValidationError("filters.store_ids must be a non-empty string array")
            placeholders = ",".join("?" for _ in store_ids)
            clauses.append(f"m.store_id IN ({placeholders})")
            parameters.extend(store_ids)

        folder_prefix = filters.get("folder_path_prefix")
        if folder_prefix is not None:
            if not isinstance(folder_prefix, str):
                raise ValidationError("filters.folder_path_prefix must be a string")
            clauses.append("m.folder_path LIKE ? ESCAPE '\\'")
            parameters.append(_escape_like(folder_prefix) + "%")

        sender_email = filters.get("sender_email")
        if sender_email is not None:
            if not isinstance(sender_email, str):
                raise ValidationError("filters.sender_email must be a string")
            clauses.append("LOWER(m.sender_email) = LOWER(?)")
            parameters.append(sender_email)

        for name, operator in (("received_from", ">="), ("received_to", "<=")):
            if name not in filters:
                continue
            value = filters[name]
            if not isinstance(value, str) or not value:
                raise ValidationError(
                    f"filters.{name} must be an ISO-8601 timestamp with a UTC offset"
                )
            try:
                canonical = canonical_utc_timestamp(value)
            except ValueError as exc:
                raise ValidationError(
                    f"filters.{name} must be an ISO-8601 timestamp with a UTC offset"
                ) from exc
            clauses.append(f"m.received_at {operator} ?")
            parameters.append(canonical)

        has_attachments = filters.get("has_attachments")
        if has_attachments is not None:
            if not isinstance(has_attachments, bool):
                raise ValidationError("filters.has_attachments must be boolean")
            predicate = "EXISTS" if has_attachments else "NOT EXISTS"
            clauses.append(
                f"{predicate} (SELECT 1 FROM attachments a WHERE a.message_id = m.id)"
            )

        return (" AND " + " AND ".join(clauses)) if clauses else "", parameters

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
            "message_id": row["message_id"],
            "text": row["text"],
            "source_kind": row["source_kind"],
            "source_label": row["source_label"],
            "entry_id": row["entry_id"],
            "store_id": row["store_id"],
            "folder_entry_id": row["folder_entry_id"],
            "subject": row["subject"],
            "sender_name": row["sender_name"],
            "sender_email": row["sender_email"],
            "received_at": row["received_at"],
            "folder_path": row["folder_path"],
            "store_name": row["store_name"],
            "internet_message_id": row["internet_message_id"],
            "conversation_id": row["conversation_id"],
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
        filters: object,
        embedder: Embedder,
        candidate_limit: int,
    ) -> list[dict[str, Any]]:
        filter_sql, filter_parameters = self._filters(filters)
        query_vector = embedder.embed_many([query])[0]
        candidates: dict[int, dict[str, Any]] = {}
        select_fields = """
            c.id AS chunk_id, c.message_id, c.text, c.source_kind, c.source_label,
            m.entry_id, m.store_id, m.folder_entry_id, m.subject,
            m.sender_name, m.sender_email, m.received_at, m.folder_path,
            m.store_name, m.internet_message_id, m.conversation_id
        """

        with self.session() as connection:
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
                    JOIN messages m ON m.id = c.message_id
                    WHERE {table_name} MATCH ? {filter_sql}
                    ORDER BY lexical_rank
                    LIMIT ?
                """
                parameters = [expression, *filter_parameters, candidate_limit]
                rows = list(connection.execute(lexical_sql, parameters))
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
                JOIN messages m ON m.id = c.message_id
                WHERE c.embedding_model = ? AND c.embedding_dim = ?
                  AND c.embedding_fingerprint = ? {filter_sql}
            """
            vector_parameters = [
                embedder.name,
                embedder.dimensions,
                embedder.fingerprint,
                *filter_parameters,
            ]
            best_vector_chunk_by_message: dict[
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
                message_id = int(row["message_id"])
                current = best_vector_chunk_by_message.get(message_id)
                if current is None or item[:2] > current[:2]:
                    best_vector_chunk_by_message[message_id] = item

            vector_candidates = heapq.nlargest(
                candidate_limit,
                best_vector_chunk_by_message.values(),
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

        folded_query = query.casefold()
        ranked_chunks: list[tuple[float, dict[str, Any]]] = []
        for candidate in candidates.values():
            score = 0.62 * candidate["vector"] + 0.38 * candidate["lexical"]
            metadata = " ".join(
                str(candidate[field])
                for field in ("subject", "sender_name", "sender_email", "folder_path")
                if candidate.get(field)
            ).casefold()
            if folded_query in metadata:
                score += 0.12
            ranked_chunks.append((min(1.0, score), candidate))
        ranked_chunks.sort(key=lambda item: (-item[0], item[1]["message_id"], item[1]["chunk_id"]))

        by_message: dict[int, dict[str, Any]] = {}
        for score, candidate in ranked_chunks:
            message_id = int(candidate["message_id"])
            source = candidate["source_label"]
            result = by_message.get(message_id)
            if result is None:
                result = {
                    "entry_id": candidate["entry_id"],
                    "store_id": candidate["store_id"],
                    "folder_entry_id": candidate["folder_entry_id"],
                    "subject": candidate["subject"],
                    "sender_name": candidate["sender_name"],
                    "sender_email": candidate["sender_email"],
                    "received_at": candidate["received_at"],
                    "folder_path": candidate["folder_path"],
                    "store_name": candidate["store_name"],
                    "internet_message_id": candidate["internet_message_id"],
                    "conversation_id": candidate["conversation_id"],
                    "vector_similarity": float(candidate["vector"]),
                    "vector_available": bool(candidate.get("vector_available", False)),
                    "lexical_score": float(candidate["lexical"]),
                    "lexical_match_kind": str(
                        candidate.get("lexical_match_kind", "")
                    ),
                    "hybrid_score": float(score),
                    "snippet": self._snippet(candidate["text"], query),
                    "matched_sources": [source],
                }
                by_message[message_id] = result
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
                result["hybrid_score"] = max(
                    float(result["hybrid_score"]),
                    float(score),
                )
                if source not in result["matched_sources"] and len(result["matched_sources"]) < 8:
                    result["matched_sources"].append(source)

        results = list(by_message.values())
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
                str(result["store_id"]),
                str(result["entry_id"]),
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
                str(result["store_id"]),
                str(result["entry_id"]),
            )
        )

        # Keep independently bounded vector and lexical pools, then merge by
        # Outlook identity. A current-model literal hit can have a weak cosine
        # score (especially for a short query), so it must not be displaced by
        # the vector top-K before the service-layer literal gate sees it.
        merged: list[dict[str, Any]] = []
        seen_identities: set[tuple[str, str]] = set()
        for result in vector_results[:limit] + lexical_results[:limit]:
            identity = (str(result["store_id"]), str(result["entry_id"]))
            if identity in seen_identities:
                continue
            seen_identities.add(identity)
            merged.append(result)
        return merged

    def stats(self) -> dict[str, Any]:
        with self.session() as connection:
            counts = {
                table: int(connection.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0])
                for table in ("messages", "attachments", "chunks")
            }
            extraction = {
                row["extraction_status"]: int(row["amount"])
                for row in connection.execute(
                    """
                    SELECT extraction_status, COUNT(*) AS amount
                    FROM attachments GROUP BY extraction_status
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
            "attachment_extraction": extraction,
            "database_bytes": database_bytes,
            "schema_version": SCHEMA_VERSION,
            "embedding_model": meta.get("embedding_model"),
            "embedding_dim": int(meta["embedding_dim"]) if "embedding_dim" in meta else None,
            "embedding_fingerprint": meta.get("embedding_fingerprint"),
        }
