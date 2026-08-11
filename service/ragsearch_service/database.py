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


SCHEMA_VERSION = 1
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
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.session() as connection:
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
                    UNIQUE (message_id, source_key, ordinal)
                );

                CREATE INDEX IF NOT EXISTS idx_chunks_message
                    ON chunks (message_id);
                CREATE INDEX IF NOT EXISTS idx_chunks_embedding
                    ON chunks (embedding_model, embedding_dim);

                CREATE VIRTUAL TABLE IF NOT EXISTS chunks_fts USING fts5(
                    text,
                    content='chunks',
                    content_rowid='id',
                    tokenize='unicode61 remove_diacritics 2'
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
                        ordinal, text, text_hash, embedding, embedding_model, embedding_dim
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                    ),
                )

            connection.execute(
                "INSERT OR REPLACE INTO service_meta(key, value) VALUES ('embedding_model', ?)",
                (embedder.name,),
            )
            connection.execute(
                "INSERT OR REPLACE INTO service_meta(key, value) VALUES ('embedding_dim', ?)",
                (str(embedder.dimensions),),
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

        received_from = filters.get("received_from")
        if received_from is not None:
            if not isinstance(received_from, str):
                raise ValidationError("filters.received_from must be a string")
            clauses.append("m.received_at >= ?")
            parameters.append(received_from)
        received_to = filters.get("received_to")
        if received_to is not None:
            if not isinstance(received_to, str):
                raise ValidationError("filters.received_to must be a string")
            clauses.append("m.received_at <= ?")
            parameters.append(received_to)

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
            lexical_rows: list[tuple[sqlite3.Row, float]] = []
            if fts_expression:
                lexical_sql = f"""
                    SELECT {select_fields}, c.embedding, c.embedding_model, c.embedding_dim,
                           bm25(chunks_fts) AS lexical_rank
                    FROM chunks_fts
                    JOIN chunks c ON c.id = chunks_fts.rowid
                    JOIN messages m ON m.id = c.message_id
                    WHERE chunks_fts MATCH ? {filter_sql}
                    ORDER BY lexical_rank
                    LIMIT ?
                """
                parameters = [fts_expression, *filter_parameters, candidate_limit]
                for row in connection.execute(lexical_sql, parameters):
                    lexical_rows.append((row, max(0.0, -float(row["lexical_rank"]))))

            maximum_lexical = max((score for _, score in lexical_rows), default=1.0) or 1.0
            for row, raw_score in lexical_rows:
                payload = self._row_payload(row)
                payload["lexical"] = raw_score / maximum_lexical
                payload["vector"] = 0.0
                try:
                    stored = blob_to_vector(row["embedding"], int(row["embedding_dim"]))
                    if row["embedding_model"] == embedder.name and len(stored) == len(query_vector):
                        payload["vector"] = max(
                            0.0, cosine_for_normalized(query_vector, stored)
                        )
                except (TypeError, ValueError):
                    pass
                candidates[int(row["chunk_id"])] = payload

            vector_sql = f"""
                SELECT {select_fields}, c.embedding
                FROM chunks c
                JOIN messages m ON m.id = c.message_id
                WHERE c.embedding_model = ? AND c.embedding_dim = ? {filter_sql}
            """
            vector_parameters = [embedder.name, embedder.dimensions, *filter_parameters]
            heap: list[tuple[float, int, sqlite3.Row]] = []
            for row in connection.execute(vector_sql, vector_parameters):
                try:
                    stored = blob_to_vector(row["embedding"], embedder.dimensions)
                except (TypeError, ValueError):
                    continue
                similarity = max(0.0, cosine_for_normalized(query_vector, stored))
                item = (similarity, int(row["chunk_id"]), row)
                if len(heap) < candidate_limit:
                    heapq.heappush(heap, item)
                elif item[:2] > heap[0][:2]:
                    heapq.heapreplace(heap, item)

            for similarity, chunk_id, row in heap:
                existing = candidates.get(chunk_id)
                if existing is None:
                    existing = self._row_payload(row)
                    existing["lexical"] = 0.0
                    candidates[chunk_id] = existing
                existing["vector"] = max(float(existing.get("vector", 0.0)), similarity)

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
                    "score": round(score, 6),
                    "snippet": self._snippet(candidate["text"], query),
                    "matched_sources": [source],
                }
                by_message[message_id] = result
            elif source not in result["matched_sources"] and len(result["matched_sources"]) < 8:
                result["matched_sources"].append(source)

        results = sorted(by_message.values(), key=lambda result: -float(result["score"]))
        return results[:limit]

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
        }
