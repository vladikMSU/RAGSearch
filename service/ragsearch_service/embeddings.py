from __future__ import annotations

import hashlib
import importlib.metadata as importlib_metadata
import math
import re
import threading
import unicodedata
from array import array
from collections.abc import Iterable, Sequence
from pathlib import Path
from typing import Protocol


_TOKEN = re.compile(r"[\w@.+-]+", re.UNICODE)


class Embedder(Protocol):
    name: str
    dimensions: int
    fingerprint: str

    def embed_many(self, texts: Sequence[str]) -> list[list[float]]: ...


class HashingEmbedder:
    """Small deterministic feature-hashing embedder with no model downloads."""

    def __init__(self, dimensions: int = 256) -> None:
        if dimensions < 32:
            raise ValueError("dimensions must be at least 32")
        self.dimensions = dimensions
        self.name = f"hashing-v1-{dimensions}"
        self.fingerprint = _contract_fingerprint(
            "provider=ragsearch-hashing",
            "algorithm_revision=1",
            f"dimensions={dimensions}",
            f"token_pattern={_TOKEN.pattern}",
            f"unicode_version={unicodedata.unidata_version}",
            "digest=blake2b-64-little-endian",
            "features=word:1.0,trigram:0.20,bigram:0.60",
            "normalization=l2",
        )

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

        artifact_root = _resolve_model_artifact_root(model_name)
        artifact_manifest = _model_artifact_manifest(artifact_root)
        self._model = SentenceTransformer(
            str(artifact_root),
            local_files_only=True,
        )
        self.dimensions = int(self._model.get_embedding_dimension())
        self.name = f"sentence-transformers:{model_name}"
        artifact_fingerprint = _fingerprint_model_directory(
            artifact_root,
            artifact_manifest,
        )
        self.fingerprint = _contract_fingerprint(
            "provider=sentence-transformers",
            f"sentence-transformers={_package_version('sentence-transformers')}",
            f"transformers={_package_version('transformers')}",
            f"tokenizers={_package_version('tokenizers')}",
            f"torch={_package_version('torch')}",
            f"artifacts={artifact_fingerprint}",
        )
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
    if provider == "hash":
        if model_name is not None:
            raise ValueError("model_name is only valid for sentence-transformers")
        return HashingEmbedder()
    if provider == "sentence-transformers":
        return SentenceTransformersEmbedder(model_name or "")
    raise ValueError(f"Unknown embedding provider: {provider}")


def _contract_fingerprint(*parts: str) -> str:
    digest = hashlib.sha256()
    for part in parts:
        encoded = part.encode("utf-8")
        digest.update(len(encoded).to_bytes(8, "big"))
        digest.update(encoded)
    return "sha256:" + digest.hexdigest()


def _package_version(distribution: str) -> str:
    try:
        return importlib_metadata.version(distribution)
    except importlib_metadata.PackageNotFoundError as exc:
        raise RuntimeError(
            f"Cannot fingerprint embedding runtime: missing package metadata for {distribution}"
        ) from exc


def _resolve_model_artifact_root(model_name: str) -> Path:
    configured = Path(model_name).expanduser()
    if configured.exists():
        if not configured.is_dir():
            raise ValueError("sentence-transformers model path must be a directory")
        return configured.resolve(strict=True)

    try:
        from huggingface_hub import snapshot_download
    except ImportError as exc:
        raise RuntimeError(
            "huggingface-hub is required to resolve a cached model identifier"
        ) from exc
    try:
        resolved = Path(snapshot_download(repo_id=model_name, local_files_only=True))
    except Exception as exc:
        raise RuntimeError(
            f"Cannot resolve locally cached sentence-transformers model {model_name!r}"
        ) from exc
    if not resolved.is_dir():
        raise RuntimeError(
            f"Locally cached sentence-transformers model {model_name!r} is not a directory"
        )
    return resolved.resolve(strict=True)


def _model_artifact_manifest(root: Path) -> tuple[tuple[str, int, int], ...]:
    root = Path(root).resolve(strict=True)
    if not root.is_dir():
        raise ValueError("model artifact root must be a directory")

    manifest = tuple(
        (
            path.relative_to(root).as_posix(),
            path.stat().st_size,
            path.stat().st_mtime_ns,
        )
        for path in sorted(
            (path for path in root.rglob("*") if path.is_file()),
            key=lambda path: path.relative_to(root).as_posix(),
        )
    )
    if not manifest:
        raise ValueError("model artifact directory must contain at least one file")
    return manifest


def _fingerprint_model_directory(
    root: Path,
    manifest: tuple[tuple[str, int, int], ...] | None = None,
) -> str:
    root = Path(root).resolve(strict=True)
    manifest = manifest or _model_artifact_manifest(root)

    digest = hashlib.sha256()
    for relative_name, expected_size, expected_mtime_ns in manifest:
        relative = relative_name.encode("utf-8")
        digest.update(len(relative).to_bytes(8, "big"))
        digest.update(relative)
        path = root / Path(relative_name)
        with path.open("rb") as stream:
            while chunk := stream.read(1024 * 1024):
                digest.update(chunk)
        stat = path.stat()
        if stat.st_size != expected_size or stat.st_mtime_ns != expected_mtime_ns:
            raise RuntimeError(
                f"Model artifact changed while its fingerprint was being computed: {relative_name}"
            )
    return "sha256:" + digest.hexdigest()


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
