from __future__ import annotations

import argparse
import logging
from dataclasses import replace
from pathlib import Path

from .app import SearchService
from .config import Settings
from .http_api import create_http_server


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Local document hybrid-search service")
    parser.add_argument("--port", type=int, default=8765, help="Loopback TCP port (default: 8765)")
    parser.add_argument("--data-dir", type=Path, help="Override the local data directory")
    parser.add_argument("--token-path", type=Path, help="Override the token file path")
    parser.add_argument(
        "--embedding",
        choices=("hash", "sentence-transformers"),
        default="hash",
        help="Embedding provider (default: deterministic dependency-free hash)",
    )
    parser.add_argument("--model", help="Locally cached sentence-transformers model name/path")
    parser.add_argument("--verbose", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    logging.basicConfig(
        level=logging.INFO if args.verbose else logging.WARNING,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )

    if not 1 <= args.port <= 65535:
        raise SystemExit("--port must be between 1 and 65535")
    if args.embedding == "sentence-transformers" and not args.model:
        raise SystemExit("--model is required with --embedding sentence-transformers")

    if args.data_dir:
        settings = Settings.explicit(
            args.data_dir,
            token_path=args.token_path,
            port=args.port,
        )
    else:
        settings = Settings.default(port=args.port)
        settings = replace(
            settings,
            token_path=args.token_path or settings.token_path,
        )

    service = SearchService(
        settings,
        embedding_provider=args.embedding,
        embedding_model=args.model,
    )
    server = create_http_server(service)
    print(f"RAGSearch service listening on http://{settings.host}:{server.server_port}")
    print(f"Token file: {settings.token_path}")
    print(f"Database: {settings.database_path}")
    try:
        server.serve_forever(poll_interval=0.25)
    except KeyboardInterrupt:
        print("Stopping RAGSearch service")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
