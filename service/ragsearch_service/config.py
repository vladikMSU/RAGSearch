from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


def _default_local_app_data() -> Path:
    configured = os.environ.get("LOCALAPPDATA")
    if configured:
        return Path(configured)
    return Path.home() / "AppData" / "Local"


@dataclass(frozen=True)
class Settings:
    data_dir: Path
    database_path: Path
    spool_dir: Path
    token_path: Path
    host: str = "127.0.0.1"
    port: int = 8765
    request_limit_bytes: int = 32 * 1024 * 1024
    attachment_limit_bytes: int = 64 * 1024 * 1024
    chunk_chars: int = 1_200
    chunk_overlap_chars: int = 180
    vector_candidate_limit: int = 200
    delete_spool_after_ingest: bool = False

    @classmethod
    def default(cls, *, port: int = 8765) -> "Settings":
        data_dir = _default_local_app_data() / "RAGSearch"
        return cls(
            data_dir=data_dir,
            database_path=data_dir / "ragsearch.sqlite3",
            spool_dir=data_dir / "spool",
            token_path=data_dir / "service-token",
            port=port,
        )

    @classmethod
    def explicit(
        cls,
        data_dir: Path,
        *,
        spool_dir: Path | None = None,
        token_path: Path | None = None,
        port: int = 0,
        delete_spool_after_ingest: bool = False,
    ) -> "Settings":
        data_dir = Path(data_dir)
        return cls(
            data_dir=data_dir,
            database_path=data_dir / "ragsearch.sqlite3",
            spool_dir=Path(spool_dir) if spool_dir else data_dir / "spool",
            token_path=Path(token_path) if token_path else data_dir / "service-token",
            port=port,
            delete_spool_after_ingest=delete_spool_after_ingest,
        )

    def ensure_directories(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.spool_dir.mkdir(parents=True, exist_ok=True)
        self.token_path.parent.mkdir(parents=True, exist_ok=True)
