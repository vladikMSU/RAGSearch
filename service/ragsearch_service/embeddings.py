from __future__ import annotations

import hashlib
import math
import re
import threading
from array import array
from collections.abc import Iterable, Sequence
from typing import Protocol


_TOKEN = re.compile(r"[\w@.+-]+", re.UNICODE)


class Embedder(Protocol):
    name: str
    dimensions: int

    def embed_many(self, texts: Sequence[str]) -> list[list[float]]: ...


class HashingEmbedder:
    """Small deterministic feature-hashing embedder with no model downloads."""

    def __init__(self, dimensions: int = 256) -> None:
        if dimensions < 32:
            raise ValueError("dimensions must be at least 32")
        self.dimensions = dimensions
        self.name = f"hashing-v1-{dimensions}"

    def _add(self, vector: list[float], feature: str, weight: float) -> None:
        digest = hashlib.blake2b(feature.encode("utf-8"), digest_size=8).digest()
        number = int.from_bytes(digest, "little")
        index = number % self.dimensions
        sign = -1.0 if number & (1 << 63) else 1.0
        vector[index] += sign * weight

    def embed(self, text: str) -> list[float]:
        tokens = [match.group(0).casefold() for match in _TOKEN.finditer(text)]
        vector = [0.0] * self.dimensions
        for token in tokens:
            self._add(vector, "w:" + token, 1.0)
            if len(token) >= 5:
                padded = "^" + token + "$"
                for offset in range(len(padded) - 2):
                    self._add(vector, "c:" + padded[offset : offset + 3], 0.20)
        for left, right in zip(tokens, tokens[1:]):
            self._add(vector, f"b:{left}\u241f{right}", 0.60)

        magnitude = math.sqrt(sum(component * component for component in vector))
        if magnitude:
            vector = [component / magnitude for component in vector]
        return vector

    def embed_many(self, texts: Sequence[str]) -> list[list[float]]:
        return [self.embed(text) for text in texts]


class SentenceTransformersEmbedder:
    """Optional provider. Models must already exist in the local cache."""

    def __init__(self, model_name: str) -> None:
        if not model_name:
            raise ValueError("model_name is required for sentence-transformers")
        try:
            from sentence_transformers import SentenceTransformer
        except ImportError as exc:
            raise RuntimeError(
                "sentence-transformers is not installed; the default hashing provider needs no dependency"
            ) from exc

        try:
            self._model = SentenceTransformer(model_name, local_files_only=True)
        except TypeError as exc:
            raise RuntimeError(
                "Installed sentence-transformers does not support local_files_only; upgrade it to avoid downloads"
            ) from exc
        get_dimensions = getattr(self._model, "get_embedding_dimension", None)
        if get_dimensions is None:
            get_dimensions = self._model.get_sentence_embedding_dimension
        self.dimensions = int(get_dimensions())
        self.name = f"sentence-transformers:{model_name}"
        self._lock = threading.Lock()

    def embed_many(self, texts: Sequence[str]) -> list[list[float]]:
        if not texts:
            return []
        with self._lock:
            encoded = self._model.encode(
                list(texts),
                normalize_embeddings=True,
                show_progress_bar=False,
            )
        return [[float(value) for value in row] for row in encoded]


def create_embedder(provider: str = "hash", model_name: str | None = None) -> Embedder:
    normalized = provider.strip().casefold()
    if normalized in {"hash", "hashing", "fallback"}:
        return HashingEmbedder()
    if normalized in {"sentence-transformers", "sentence_transformers", "st"}:
        return SentenceTransformersEmbedder(model_name or "")
    raise ValueError(f"Unknown embedding provider: {provider}")


def vector_to_blob(values: Iterable[float]) -> bytes:
    return array("f", values).tobytes()


def blob_to_vector(blob: bytes, dimensions: int) -> array:
    values = array("f")
    values.frombytes(blob)
    if len(values) != dimensions:
        raise ValueError(f"Invalid vector length: expected {dimensions}, got {len(values)}")
    return values


def cosine_for_normalized(left: Sequence[float], right: Sequence[float]) -> float:
    return float(sum(a * b for a, b in zip(left, right)))
