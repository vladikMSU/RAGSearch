from __future__ import annotations

import hashlib
from typing import Any

from .attachments import ExtractedAttachment, extract_attachment
from .chunking import chunk_text, normalize_text
from .config import Settings
from .database import Database
from .embeddings import Embedder, create_embedder
from .errors import ValidationError
from .security import ensure_private_path, load_or_create_token


def _required_text(payload: dict[str, Any], name: str) -> str:
    value = payload.get(name)
    if not isinstance(value, str) or not value.strip():
        raise ValidationError(f"{name} must be a non-empty string")
    return value.strip()


def _optional_text(payload: dict[str, Any], name: str) -> str:
    value = payload.get(name)
    if value is None:
        return ""
    if not isinstance(value, str):
        raise ValidationError(f"{name} must be a string or null")
    return value


def _recipients(payload: dict[str, Any], name: str) -> str:
    value = payload.get(name)
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, list) and all(isinstance(item, str) for item in value):
        return "; ".join(item for item in value if item)
    raise ValidationError(f"{name} must be a string, string array, or null")


def _optional_timestamp(payload: dict[str, Any], name: str) -> str | None:
    value = payload.get(name)
    if value in {None, ""}:
        return None
    if not isinstance(value, str):
        raise ValidationError(f"{name} must be an ISO-8601 string or null")
    return value


class SearchService:
    def __init__(
        self,
        settings: Settings,
        *,
        embedder: Embedder | None = None,
        embedding_provider: str = "hash",
        embedding_model: str | None = None,
    ) -> None:
        self.settings = settings
        self.settings.ensure_directories()
        ensure_private_path(settings.data_dir)
        ensure_private_path(settings.spool_dir)
        if settings.token_path.parent not in {settings.data_dir, settings.spool_dir}:
            ensure_private_path(settings.token_path.parent)
        self.token = load_or_create_token(settings.token_path)
        self.embedder = embedder or create_embedder(embedding_provider, embedding_model)
        self.database = Database(settings.database_path)
        self.database.initialize()
        ensure_private_path(settings.database_path)

    def health(self) -> dict[str, object]:
        return {"status": "ok" if self.database.ping() else "error"}

    def _normalize_message(self, raw: object) -> dict[str, Any]:
        if not isinstance(raw, dict):
            raise ValidationError("Each messages item must be an object")
        body = _optional_text(raw, "body")
        attachments = raw.get("attachments", [])
        if attachments is None:
            attachments = []
        if not isinstance(attachments, list):
            raise ValidationError("attachments must be an array")
        return {
            "entry_id": _required_text(raw, "entry_id"),
            "store_id": _required_text(raw, "store_id"),
            "folder_entry_id": _required_text(raw, "folder_entry_id"),
            "folder_path": _optional_text(raw, "folder_path"),
            "store_name": _optional_text(raw, "store_name"),
            "subject": _optional_text(raw, "subject"),
            "sender_name": _optional_text(raw, "sender_name"),
            "sender_email": _optional_text(raw, "sender_email"),
            "to": _recipients(raw, "to"),
            "cc": _recipients(raw, "cc"),
            "sent_at": _optional_timestamp(raw, "sent_at"),
            "received_at": _optional_timestamp(raw, "received_at"),
            "modified_at": _optional_timestamp(raw, "modified_at"),
            "internet_message_id": _optional_text(raw, "internet_message_id"),
            "conversation_id": _optional_text(raw, "conversation_id"),
            "body": body,
            "body_hash": hashlib.sha256(body.encode("utf-8")).hexdigest() if body else "",
            "attachments": attachments,
        }

    def _metadata_text(self, message: dict[str, Any]) -> str:
        fields = (
            ("Subject", message["subject"]),
            ("From", " ".join(filter(None, (message["sender_name"], message["sender_email"])))),
            ("To", message["to"]),
            ("Cc", message["cc"]),
            ("Folder", message["folder_path"]),
            ("Store", message["store_name"]),
            ("Conversation", message["conversation_id"]),
            ("Internet-Message-ID", message["internet_message_id"]),
        )
        return "\n".join(f"{label}: {value}" for label, value in fields if value)

    def _prepare_attachment(
        self,
        raw: object,
        attachment_index: int,
    ) -> tuple[dict[str, Any], ExtractedAttachment]:
        if not isinstance(raw, dict):
            raise ValidationError(f"attachments[{attachment_index}] must be an object")
        name = _optional_text(raw, "name")
        content_type = _optional_text(raw, "content_type")
        size = raw.get("size", 0)
        if size is None:
            size = 0
        if not isinstance(size, int) or isinstance(size, bool) or size < 0:
            raise ValidationError(f"attachments[{attachment_index}].size must be a non-negative integer")
        extracted = extract_attachment(
            raw,
            spool_dir=self.settings.spool_dir,
            limit_bytes=self.settings.attachment_limit_bytes,
        )
        record = {
            "attachment_index": attachment_index,
            "name": name,
            "size": size,
            "content_type": content_type,
            "extraction_status": extracted.status,
            "extraction_error": extracted.error,
            "text_hash": (
                hashlib.sha256(extracted.text.encode("utf-8")).hexdigest()
                if extracted.text
                else ""
            ),
        }
        return record, extracted

    def _embed_chunks(self, chunks: list[dict[str, Any]]) -> None:
        batch_size = 64
        for offset in range(0, len(chunks), batch_size):
            batch = chunks[offset : offset + batch_size]
            embeddings = self.embedder.embed_many([chunk["text"] for chunk in batch])
            if len(embeddings) != len(batch):
                raise RuntimeError("Embedding provider returned an invalid batch length")
            for chunk, embedding in zip(batch, embeddings):
                if len(embedding) != self.embedder.dimensions:
                    raise RuntimeError("Embedding provider returned an invalid vector dimension")
                chunk["embedding"] = embedding

    def _ingest_one(self, raw: object) -> None:
        message = self._normalize_message(raw)
        attachment_records: list[dict[str, Any]] = []
        extracted_attachments: list[tuple[dict[str, Any], ExtractedAttachment]] = []
        for index, attachment in enumerate(message.pop("attachments")):
            record, extracted = self._prepare_attachment(attachment, index)
            attachment_records.append(record)
            extracted_attachments.append((record, extracted))

        chunks: list[dict[str, Any]] = []
        metadata = normalize_text(self._metadata_text(message))
        if metadata:
            chunks.append(
                {
                    "source_key": "metadata",
                    "source_kind": "metadata",
                    "source_label": "message_metadata",
                    "ordinal": 0,
                    "text": metadata,
                }
            )
        for ordinal, text in enumerate(
            chunk_text(
                message.pop("body"),
                max_chars=self.settings.chunk_chars,
                overlap_chars=self.settings.chunk_overlap_chars,
            )
        ):
            chunks.append(
                {
                    "source_key": "body",
                    "source_kind": "body",
                    "source_label": "body",
                    "ordinal": ordinal,
                    "text": text,
                }
            )
        for record, extracted in extracted_attachments:
            for ordinal, text in enumerate(
                chunk_text(
                    extracted.text,
                    max_chars=self.settings.chunk_chars,
                    overlap_chars=self.settings.chunk_overlap_chars,
                )
            ):
                chunks.append(
                    {
                        "attachment_index": record["attachment_index"],
                        "source_key": f"attachment:{record['attachment_index']}",
                        "source_kind": "attachment",
                        "source_label": f"attachment:{record['name'] or record['attachment_index']}",
                        "ordinal": ordinal,
                        "text": text,
                    }
                )

        self._embed_chunks(chunks)
        self.database.upsert_message(
            message,
            attachment_records,
            chunks,
            embedder=self.embedder,
        )

        if self.settings.delete_spool_after_ingest:
            for _, extracted in extracted_attachments:
                if extracted.safe_path is not None:
                    try:
                        extracted.safe_path.unlink()
                    except OSError:
                        pass

    def ingest_messages(self, payload: object) -> dict[str, Any]:
        if not isinstance(payload, dict):
            raise ValidationError("JSON body must be an object")
        messages = payload.get("messages")
        if not isinstance(messages, list):
            raise ValidationError("messages must be an array")

        accepted = 0
        errors: list[dict[str, Any]] = []
        for index, raw in enumerate(messages):
            try:
                self._ingest_one(raw)
                accepted += 1
            except Exception as exc:
                error: dict[str, Any] = {"index": index, "error": str(exc)}
                if isinstance(raw, dict):
                    if isinstance(raw.get("entry_id"), str):
                        error["entry_id"] = raw["entry_id"]
                    if isinstance(raw.get("store_id"), str):
                        error["store_id"] = raw["store_id"]
                errors.append(error)
        return {"accepted": accepted, "failed": len(errors), "errors": errors}

    def search(self, payload: object) -> dict[str, Any]:
        if not isinstance(payload, dict):
            raise ValidationError("JSON body must be an object")
        query = payload.get("query")
        if not isinstance(query, str) or not query.strip():
            raise ValidationError("query must be a non-empty string")
        limit = payload.get("limit", 20)
        if not isinstance(limit, int) or isinstance(limit, bool) or not 1 <= limit <= 100:
            raise ValidationError("limit must be an integer between 1 and 100")
        results = self.database.search(
            query.strip(),
            limit=limit,
            filters=payload.get("filters"),
            embedder=self.embedder,
            candidate_limit=max(limit * 10, self.settings.vector_candidate_limit),
        )
        return {"results": results}

    def clear_index(self) -> dict[str, int]:
        return self.database.clear_index()

    def stats(self) -> dict[str, Any]:
        return self.database.stats()
