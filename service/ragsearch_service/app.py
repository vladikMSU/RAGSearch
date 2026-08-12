from __future__ import annotations

import hashlib
import re
from typing import Any

from .attachments import ExtractedAttachment, extract_attachment
from .chunking import chunk_text, normalize_text
from .config import Settings
from .database import Database
from .embeddings import Embedder, create_embedder
from .errors import ValidationError
from .security import ensure_private_path, load_or_create_token
from .timestamps import canonical_utc_timestamp


_SEARCH_TOKEN = re.compile(r"[\w@.+-]+", re.UNICODE)


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
    raise ValidationError(f"{name} must be a string or null")


def _optional_timestamp(payload: dict[str, Any], name: str) -> str | None:
    value = payload.get(name)
    if value is None:
        return None
    if not isinstance(value, str) or not value:
        raise ValidationError(
            f"{name} must be an ISO-8601 timestamp with a UTC offset or null"
        )
    try:
        return canonical_utc_timestamp(value)
    except ValueError as exc:
        raise ValidationError(
            f"{name} must be an ISO-8601 timestamp with a UTC offset or null"
        ) from exc


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
        self.database.validate_embedding_model(self.embedder)
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
        candidate_message_limit = max(
            limit,
            self.settings.search_candidate_message_limit,
        )
        candidates = self.database.search(
            query.strip(),
            limit=candidate_message_limit,
            filters=payload.get("filters"),
            embedder=self.embedder,
            candidate_limit=max(
                candidate_message_limit * 10,
                self.settings.vector_candidate_limit,
            ),
        )

        minimum_similarity = self.settings.minimum_vector_similarity
        if self.embedder.name.startswith("hashing-"):
            # Cosine distributions are model-specific. The dependency-free
            # The hashing provider needs a lower floor than a dense semantic model.
            minimum_similarity = self.settings.hashing_minimum_vector_similarity

        vector_candidates = [
            item for item in candidates if bool(item["vector_available"])
        ]
        best_similarity = max(
            (float(item["vector_similarity"]) for item in vector_candidates),
            default=0.0,
        )
        cutoff_similarity = max(
            minimum_similarity,
            best_similarity - self.settings.vector_similarity_window,
        )
        vector_eligible = [
            item
            for item in vector_candidates
            if float(item["vector_similarity"]) + 1e-9 >= cutoff_similarity
        ]
        lexical_matches = [
            item for item in candidates if float(item["lexical_score"]) > 0.0
        ]
        query_tokens = {
            match.group(0).casefold() for match in _SEARCH_TOKEN.finditer(query)
        }
        fragment_matches = [
            item
            for item in lexical_matches
            if item.get("lexical_match_kind") in {"prefix", "substring"}
        ]
        lexical_gate = bool(fragment_matches) or (
            len(query_tokens) == 1 and bool(lexical_matches)
        )

        lexical_matches.sort(
            key=lambda item: (
                0 if bool(item["vector_available"]) else 1,
                float(item["vector_distance"]),
                -float(item["hybrid_score"]),
                str(item["store_id"]),
                str(item["entry_id"]),
            )
        )
        lexical_identities = {
            (str(item["store_id"]), str(item["entry_id"]))
            for item in lexical_matches
        }
        semantic_matches = [
            item
            for item in vector_eligible
            if (str(item["store_id"]), str(item["entry_id"]))
            not in lexical_identities
        ]
        if lexical_gate:
            eligible = lexical_matches
        elif len(query_tokens) == 1:
            # The current multilingual paraphrase model is demonstrably
            # unreliable for isolated words (for example спорт vs куку). Do not
            # manufacture a semantic hit when there is no literal evidence.
            eligible = []
        else:
            eligible = lexical_matches + semantic_matches
        maximum_results = min(limit, self.settings.search_result_limit)
        results = eligible[:maximum_results]
        for rank, result in enumerate(results, start=1):
            result["rank"] = rank
            match_kind = str(result.get("lexical_match_kind", ""))
            result["ranking_basis"] = (
                "lexical_" + match_kind if match_kind else "vector_distance"
            )

        cutoff_similarity = min(1.0, max(0.0, cutoff_similarity))
        return {
            "results": results,
            "total": len(results),
            "mode": (
                "lexical-fragment"
                if fragment_matches
                else "lexical-exact"
                if lexical_gate
                else "single-token-no-literal"
                if len(query_tokens) == 1
                else "hybrid-semantic"
            ),
            "candidate_count": len(candidates),
            "eligible_count": len(eligible),
            "lexical_match_count": len(lexical_matches),
            "lexical_gate": lexical_gate,
            "best_vector_similarity": round(best_similarity, 6),
            "best_vector_distance": round(1.0 - best_similarity, 6),
            "cutoff_similarity": round(cutoff_similarity, 6),
            "cutoff_distance": round(1.0 - cutoff_similarity, 6),
            "max_results": maximum_results,
            "ranking": "lexical_gate_then_vector_distance_asc",
        }

    def clear_index(self) -> dict[str, int]:
        return self.database.clear_index()

    def stats(self) -> dict[str, Any]:
        return self.database.stats()
