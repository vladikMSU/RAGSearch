from __future__ import annotations

import base64
import binascii
import hashlib
import json
import math
import re
from typing import Any

from .attachments import ExtractedPart, extract_part
from .chunking import chunk_text, normalize_text
from .config import Settings
from .database import Database
from .embeddings import Embedder, create_embedder
from .errors import ValidationError
from .security import ensure_private_path, load_or_create_token


_SEARCH_TOKEN = re.compile(r"[\w@.+-]+", re.UNICODE)
PROTOCOL_VERSION = 4
_DOCUMENT_FIELDS = {"source_key", "kind", "title", "metadata", "locator", "parts"}
_PART_FIELDS = {
    "key",
    "kind",
    "name",
    "media_type",
    "size",
    "text",
    "content_base64",
    "truncated",
}


def _reject_unpaired_surrogates(value: object) -> None:
    """Reject strings that cannot be encoded as well-formed UTF-8.

    Python's JSON decoder combines a valid UTF-16 surrogate-pair escape into one
    astral code point. Any surrogate code point that remains is therefore unpaired
    and must not reach canonical JSON, hashing, SQLite, or an HTTP response.
    """

    stack = [value]
    while stack:
        current = stack.pop()
        if isinstance(current, str):
            if any(0xD800 <= ord(character) <= 0xDFFF for character in current):
                raise ValidationError(
                    "request contains an unpaired UTF-16 surrogate code point"
                )
        elif isinstance(current, dict):
            stack.extend(current.keys())
            stack.extend(current.values())
        elif isinstance(current, list):
            stack.extend(current)


def _required_text(
    payload: dict[str, Any],
    name: str,
    *,
    maximum_chars: int,
    preserve_whitespace: bool = False,
) -> str:
    value = payload.get(name)
    if not isinstance(value, str) or not value.strip():
        raise ValidationError(f"{name} must be a non-empty string")
    if not preserve_whitespace:
        value = value.strip()
    if len(value) > maximum_chars:
        raise ValidationError(f"{name} exceeds {maximum_chars} characters")
    if "\0" in value:
        raise ValidationError(f"{name} contains a null character")
    return value


def _optional_text(payload: dict[str, Any], name: str, *, maximum_chars: int) -> str:
    value = payload.get(name)
    if value is None:
        return ""
    if not isinstance(value, str):
        raise ValidationError(f"{name} must be a string or null")
    if len(value) > maximum_chars:
        raise ValidationError(f"{name} exceeds {maximum_chars} characters")
    if "\0" in value:
        raise ValidationError(f"{name} contains a null character")
    return value


def _canonical_json(
    value: object,
    name: str,
    *,
    maximum_bytes: int,
    maximum_depth: int,
    maximum_nodes: int,
) -> str:
    stack: list[tuple[object, int]] = [(value, 0)]
    nodes = 0
    while stack:
        current, depth = stack.pop()
        nodes += 1
        if nodes > maximum_nodes:
            raise ValidationError(f"{name} exceeds {maximum_nodes} JSON nodes")
        if depth > maximum_depth:
            raise ValidationError(f"{name} exceeds JSON depth {maximum_depth}")
        if isinstance(current, dict):
            if any(not isinstance(key, str) for key in current):
                raise ValidationError(f"{name} must contain valid JSON values")
            # The .NET Framework host uses DataContractJsonSerializer for the
            # envelope; this reserved key is interpreted as a runtime type hint
            # before unknown fields can be ignored. Reject it at the protocol
            # boundary so one document cannot make the whole search response
            # undecodable. All other keys remain opaque.
            if "__type" in current:
                raise ValidationError(f"{name} contains reserved JSON key __type")
            if nodes + len(stack) + len(current) > maximum_nodes:
                raise ValidationError(f"{name} exceeds {maximum_nodes} JSON nodes")
            stack.extend((item, depth + 1) for item in current.values())
        elif isinstance(current, list):
            if nodes + len(stack) + len(current) > maximum_nodes:
                raise ValidationError(f"{name} exceeds {maximum_nodes} JSON nodes")
            stack.extend((item, depth + 1) for item in current)
        elif current is not None:
            if isinstance(current, bool) or isinstance(current, str):
                continue
            if isinstance(current, (int, float)):
                try:
                    numeric = float(current)
                except OverflowError as exc:
                    raise ValidationError(
                        f"{name} contains a number outside the finite IEEE-754 range"
                    ) from exc
                if not math.isfinite(numeric):
                    raise ValidationError(
                        f"{name} contains a number outside the finite IEEE-754 range"
                    )
                continue
            raise ValidationError(f"{name} must contain valid JSON values")

    try:
        encoded = json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
        )
    except (TypeError, ValueError) as exc:
        raise ValidationError(f"{name} must contain valid JSON values") from exc
    size = len(encoded.encode("utf-8"))
    if size > maximum_bytes:
        raise ValidationError(f"{name} exceeds {maximum_bytes} UTF-8 JSON bytes")
    return encoded


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
        if settings.token_path.parent != settings.data_dir:
            ensure_private_path(settings.token_path.parent)
        self.token = load_or_create_token(settings.token_path)
        self.embedder = embedder or create_embedder(embedding_provider, embedding_model)
        self.database = Database(settings.database_path)
        self.database.initialize()
        self.database.validate_embedding_model(self.embedder)
        ensure_private_path(settings.database_path)

    def health(self) -> dict[str, object]:
        return {
            "status": "ok" if self.database.ping() else "error",
            "protocol": PROTOCOL_VERSION,
        }

    def _normalize_document(self, raw: object) -> dict[str, Any]:
        if not isinstance(raw, dict):
            raise ValidationError("JSON body must be a document object")
        unknown = set(raw) - _DOCUMENT_FIELDS
        if unknown:
            raise ValidationError(
                f"Unknown document fields: {', '.join(sorted(str(item) for item in unknown))}"
            )
        missing = _DOCUMENT_FIELDS - set(raw)
        if missing:
            raise ValidationError(
                f"Missing document fields: {', '.join(sorted(missing))}"
            )

        title = raw.get("title")
        if not isinstance(title, str):
            raise ValidationError("title must be a string")
        if len(title) > 65_536:
            raise ValidationError("title exceeds 65536 characters")
        if "\0" in title:
            raise ValidationError("title contains a null character")

        metadata = raw.get("metadata")
        if not isinstance(metadata, dict):
            raise ValidationError("metadata must be an object")
        locator = raw.get("locator")
        if not isinstance(locator, dict):
            raise ValidationError("locator must be an object")
        parts = raw.get("parts")
        if not isinstance(parts, list):
            raise ValidationError("parts must be an array")
        if len(parts) > 4_096:
            raise ValidationError("parts exceeds 4096 items")

        return {
            # Producer-owned identity is opaque. Whitespace is significant and
            # must not collapse two different source keys into one UPSERT row.
            "source_key": _required_text(
                raw,
                "source_key",
                maximum_chars=65_536,
                preserve_whitespace=True,
            ),
            "kind": _required_text(raw, "kind", maximum_chars=256),
            "title": title,
            "metadata_json": _canonical_json(
                metadata,
                "metadata",
                maximum_bytes=self.settings.metadata_json_limit_bytes,
                maximum_depth=self.settings.opaque_json_max_depth,
                maximum_nodes=self.settings.opaque_json_max_nodes,
            ),
            "locator_json": _canonical_json(
                locator,
                "locator",
                maximum_bytes=self.settings.locator_json_limit_bytes,
                maximum_depth=self.settings.opaque_json_max_depth,
                maximum_nodes=self.settings.opaque_json_max_nodes,
            ),
            "parts": parts,
        }

    @staticmethod
    def _metadata_text(
        document: dict[str, Any],
        parts: list[dict[str, Any]],
    ) -> str:
        fields = [
            f"Kind: {document['kind']}",
            f"Title: {document['title']}" if document["title"] else "",
            f"Metadata: {document['metadata_json']}"
            if document["metadata_json"] != "{}"
            else "",
        ]
        fields.extend(
            "Part: "
            + " ".join(
                value
                for value in (
                    str(part["kind"]),
                    str(part["name"]),
                    str(part["media_type"]),
                )
                if value
            )
            for part in parts
            if part["name"] or part["media_type"]
        )
        return "\n".join(field for field in fields if field)

    def _prepare_part(
        self,
        raw: object,
        part_index: int,
    ) -> tuple[dict[str, Any], ExtractedPart, bool]:
        prefix = f"parts[{part_index}]"
        if not isinstance(raw, dict):
            raise ValidationError(f"{prefix} must be an object")
        unknown = set(raw) - _PART_FIELDS
        if unknown:
            raise ValidationError(
                f"Unknown {prefix} fields: {', '.join(sorted(str(item) for item in unknown))}"
            )

        key = _required_text(raw, "key", maximum_chars=32_768)
        kind = _required_text(raw, "kind", maximum_chars=256)
        name = _optional_text(raw, "name", maximum_chars=32_768)
        media_type = _optional_text(raw, "media_type", maximum_chars=4_096)

        size = raw.get("size", 0)
        if type(size) is not int or not 0 <= size <= (1 << 63) - 1:
            raise ValidationError(f"{prefix}.size must be a bounded non-negative integer")
        truncated = raw.get("truncated", False)
        if type(truncated) is not bool:
            raise ValidationError(f"{prefix}.truncated must be boolean")

        raw_text = raw.get("text")
        if raw_text is not None and not isinstance(raw_text, str):
            raise ValidationError(f"{prefix}.text must be a string or null")
        raw_content = raw.get("content_base64")
        if raw_content is not None and not isinstance(raw_content, str):
            raise ValidationError(f"{prefix}.content_base64 must be a string or null")
        if raw_text is not None and raw_content is not None:
            raise ValidationError(
                f"{prefix} cannot contain both text and content_base64"
            )

        binary_content = raw_content is not None
        if raw_text is not None:
            extracted = ExtractedPart(raw_text, "provided_text", None)
        elif raw_content is not None:
            try:
                encoded = raw_content.encode("ascii")
                maximum_encoded = 4 * (
                    (self.settings.inline_part_limit_bytes + 2) // 3
                )
                if len(encoded) > maximum_encoded:
                    raise ValidationError(
                        f"{prefix}.content_base64 exceeds the inline part limit "
                        f"({self.settings.inline_part_limit_bytes} bytes)"
                    )
                content = base64.b64decode(encoded, validate=True)
            except ValidationError:
                raise
            except (UnicodeEncodeError, binascii.Error, ValueError) as exc:
                raise ValidationError(f"{prefix}.content_base64 is not valid base64") from exc
            if len(content) > self.settings.inline_part_limit_bytes:
                raise ValidationError(
                    f"{prefix}.content_base64 exceeds the inline part limit "
                    f"({self.settings.inline_part_limit_bytes} bytes)"
                )
            if size != len(content):
                raise ValidationError(
                    f"{prefix}.size does not match decoded content_base64 length"
                )
            extracted = extract_part(
                content,
                name=name,
                media_type=media_type,
                limit_bytes=self.settings.inline_part_limit_bytes,
                text_limit_chars=self.settings.extracted_text_limit_chars,
            )
        else:
            extracted = ExtractedPart("", "not_provided", None)

        record = {
            "part_key": key,
            "kind": kind,
            "name": name,
            "media_type": media_type,
            "size": size,
            "truncated": truncated,
            "extraction_status": extracted.status,
            "extraction_error": extracted.error,
            "text_hash": (
                hashlib.sha256(extracted.text.encode("utf-8")).hexdigest()
                if extracted.text
                else ""
            ),
        }
        return record, extracted, binary_content

    @staticmethod
    def _set_extraction_result(
        record: dict[str, Any],
        extracted: ExtractedPart,
    ) -> None:
        record["extraction_status"] = extracted.status
        record["extraction_error"] = extracted.error
        record["text_hash"] = (
            hashlib.sha256(extracted.text.encode("utf-8")).hexdigest()
            if extracted.text
            else ""
        )

    @staticmethod
    def _truncate_extraction(
        extracted: ExtractedPart,
        maximum_chars: int,
        reason: str,
    ) -> ExtractedPart:
        if len(extracted.text) <= maximum_chars:
            return extracted
        return ExtractedPart(
            extracted.text[:maximum_chars],
            "extracted_truncated",
            reason,
        )

    @staticmethod
    def _mark_index_projection_truncated(
        record: dict[str, Any],
        reason: str,
    ) -> None:
        record["extraction_status"] = "extracted_truncated"
        existing = str(record.get("extraction_error") or "")
        record["extraction_error"] = existing + ("; " if existing else "") + reason

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

    def ingest_document(self, payload: object) -> dict[str, str]:
        _reject_unpaired_surrogates(payload)
        document = self._normalize_document(payload)
        raw_parts = document.pop("parts")

        part_records: list[dict[str, Any]] = []
        extracted_parts: list[tuple[dict[str, Any], ExtractedPart, bool]] = []
        seen_part_keys: set[str] = set()
        binary_extracted_characters = 0
        for index, raw_part in enumerate(raw_parts):
            record, extracted, binary_content = self._prepare_part(raw_part, index)
            part_key = str(record["part_key"])
            if part_key in seen_part_keys:
                raise ValidationError(f"Duplicate part key: {part_key}")
            seen_part_keys.add(part_key)
            if binary_content:
                remaining = max(
                    0,
                    self.settings.binary_extracted_text_limit_chars
                    - binary_extracted_characters,
                )
                extracted = self._truncate_extraction(
                    extracted,
                    remaining,
                    "Extracted text truncated at the document aggregate binary "
                    f"limit ({self.settings.binary_extracted_text_limit_chars} characters)",
                )
                binary_extracted_characters += len(extracted.text)
                self._set_extraction_result(record, extracted)
            part_records.append(record)
            extracted_parts.append((record, extracted, binary_content))

        chunks: list[dict[str, Any]] = []
        searchable_characters = 0

        def append_producer_chunks(
            text_value: str,
            *,
            part_key: str,
            part_kind: str,
            part_label: str,
        ) -> None:
            nonlocal searchable_characters
            searchable_characters += len(text_value)
            if searchable_characters > self.settings.document_text_limit_chars:
                raise ValidationError(
                    "document searchable text exceeds "
                    f"{self.settings.document_text_limit_chars} characters"
                )
            for ordinal, text in enumerate(
                chunk_text(
                    text_value,
                    max_chars=self.settings.chunk_chars,
                    overlap_chars=self.settings.chunk_overlap_chars,
                )
            ):
                if len(chunks) >= self.settings.document_chunk_limit:
                    raise ValidationError(
                        f"document exceeds {self.settings.document_chunk_limit} chunks"
                    )
                chunks.append(
                    {
                        "part_key": part_key,
                        "part_kind": part_kind,
                        "part_label": part_label,
                        "ordinal": ordinal,
                        "text": text,
                    }
                )

        def append_derived_chunks(
            text_value: str,
            *,
            part_key: str,
            part_kind: str,
            part_label: str,
        ) -> bool:
            """Append a bounded derived projection, returning whether it was cropped."""
            nonlocal searchable_characters
            remaining_characters = max(
                0,
                self.settings.document_text_limit_chars - searchable_characters,
            )
            bounded = text_value[:remaining_characters]
            truncated = len(bounded) != len(text_value)
            searchable_characters += len(bounded)
            for ordinal, text in enumerate(
                chunk_text(
                    bounded,
                    max_chars=self.settings.chunk_chars,
                    overlap_chars=self.settings.chunk_overlap_chars,
                )
            ):
                if len(chunks) >= self.settings.document_chunk_limit:
                    truncated = True
                    break
                chunks.append(
                    {
                        "part_key": part_key,
                        "part_kind": part_kind,
                        "part_label": part_label,
                        "ordinal": ordinal,
                        "text": text,
                    }
                )
            return truncated

        # Producer-supplied text is lossless: it either fits both hard document
        # budgets or the request is rejected. Derived projections consume only the
        # remaining capacity, so metadata/attachment extraction cannot make an
        # otherwise valid producer document fail nondeterministically.
        for record, extracted, binary_content in extracted_parts:
            if not binary_content and record["extraction_status"] == "provided_text":
                label = (
                    f"attachment:{record['name'] or record['part_key']}"
                    if str(record["kind"]).casefold() == "attachment"
                    else str(record["name"] or record["part_key"])
                )
                append_producer_chunks(
                    extracted.text,
                    part_key=str(record["part_key"]),
                    part_kind=str(record["kind"]),
                    part_label=label,
                )

        metadata = normalize_text(self._metadata_text(document, part_records))
        metadata = metadata[: self.settings.metadata_search_text_limit_chars]
        append_derived_chunks(
            metadata,
            part_key="",
            part_kind="metadata",
            part_label="metadata",
        )

        for record, extracted, binary_content in extracted_parts:
            if not binary_content:
                continue
            label = (
                f"attachment:{record['name'] or record['part_key']}"
                if str(record["kind"]).casefold() == "attachment"
                else str(record["name"] or record["part_key"])
            )
            projection_truncated = append_derived_chunks(
                extracted.text,
                part_key=str(record["part_key"]),
                part_kind=str(record["kind"]),
                part_label=label,
            )
            if projection_truncated and extracted.text:
                self._mark_index_projection_truncated(
                    record,
                    "Extracted text indexing truncated by the document searchable-text "
                    "or chunk limit",
                )

        self._embed_chunks(chunks)
        self.database.upsert_document(
            document,
            part_records,
            chunks,
            embedder=self.embedder,
        )
        return {"source_key": document["source_key"], "status": "upserted"}

    def search(self, payload: object) -> dict[str, Any]:
        _reject_unpaired_surrogates(payload)
        if not isinstance(payload, dict):
            raise ValidationError("JSON body must be an object")
        unknown = set(payload) - {"query", "limit"}
        if unknown:
            raise ValidationError(
                f"Unknown search fields: {', '.join(sorted(str(item) for item in unknown))}"
            )
        query = payload.get("query")
        if not isinstance(query, str) or not query.strip():
            raise ValidationError("query must be a non-empty string")
        if len(query) > self.settings.query_limit_chars:
            raise ValidationError(
                f"query exceeds {self.settings.query_limit_chars} characters"
            )
        limit = payload.get("limit", 20)
        if not isinstance(limit, int) or isinstance(limit, bool) or not 1 <= limit <= 100:
            raise ValidationError("limit must be an integer between 1 and 100")
        candidate_document_limit = max(
            limit,
            self.settings.search_candidate_document_limit,
        )
        candidates = self.database.search(
            query.strip(),
            limit=candidate_document_limit,
            embedder=self.embedder,
            candidate_limit=max(
                candidate_document_limit * 10,
                self.settings.vector_candidate_limit,
            ),
        )

        minimum_similarity = self.settings.minimum_vector_similarity
        if self.embedder.name.startswith("hashing-"):
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
                str(item["source_key"]),
            )
        )
        lexical_identities = {str(item["source_key"]) for item in lexical_matches}
        semantic_matches = [
            item
            for item in vector_eligible
            if str(item["source_key"]) not in lexical_identities
        ]
        if lexical_gate:
            eligible = lexical_matches
        elif len(query_tokens) == 1:
            eligible = []
        else:
            eligible = lexical_matches + semantic_matches
        maximum_results = min(limit, self.settings.search_result_limit)
        results = eligible[:maximum_results]
        for rank, result in enumerate(results, start=1):
            result.pop("vector_available", None)
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
