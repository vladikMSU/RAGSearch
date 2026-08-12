from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


def _default_local_app_data() -> Path:
    configured = os.environ.get("LOCALAPPDATA")
    if configured is None or not configured.strip():
        raise RuntimeError(
            "LOCALAPPDATA is required for the default RAGSearch data directory"
        )
    return Path(configured)


@dataclass(frozen=True)
class Settings:
    data_dir: Path
    database_path: Path
    token_path: Path
    host: str = "127.0.0.1"
    port: int = 8765
    # One 8 MiB inline binary expands to roughly 10.7 MiB in base64.  The
    # remaining headroom covers text, metadata and JSON framing.
    request_limit_bytes: int = 48 * 1024 * 1024
    inline_part_limit_bytes: int = 8 * 1024 * 1024
    extracted_text_limit_chars: int = 8 * 1024 * 1024
    metadata_json_limit_bytes: int = 1 * 1024 * 1024
    locator_json_limit_bytes: int = 64 * 1024
    opaque_json_max_depth: int = 16
    opaque_json_max_nodes: int = 10_000
    document_text_limit_chars: int = 16 * 1024 * 1024
    metadata_search_text_limit_chars: int = 4 * 1024 * 1024
    binary_extracted_text_limit_chars: int = 8 * 1024 * 1024
    document_chunk_limit: int = 16_384
    query_limit_chars: int = 8_192
    chunk_chars: int = 1_200
    chunk_overlap_chars: int = 180
    vector_candidate_limit: int = 200
    search_candidate_document_limit: int = 100
    search_result_limit: int = 25
    minimum_vector_similarity: float = 0.40
    hashing_minimum_vector_similarity: float = 0.30
    vector_similarity_window: float = 0.10

    @classmethod
    def default(cls, *, port: int = 8765) -> "Settings":
        data_dir = _default_local_app_data() / "RAGSearch"
        return cls(
            data_dir=data_dir,
            database_path=data_dir / "ragsearch.sqlite3",
            token_path=data_dir / "service-token",
            port=port,
        )

    @classmethod
    def explicit(
        cls,
        data_dir: Path,
        *,
        token_path: Path | None = None,
        port: int = 0,
    ) -> "Settings":
        data_dir = Path(data_dir)
        return cls(
            data_dir=data_dir,
            database_path=data_dir / "ragsearch.sqlite3",
            token_path=Path(token_path) if token_path else data_dir / "service-token",
            port=port,
        )

    def ensure_directories(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.token_path.parent.mkdir(parents=True, exist_ok=True)
