from __future__ import annotations

import json
import logging
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any
from urllib.parse import urlsplit

from .app import PROTOCOL_VERSION, SearchService
from .errors import ValidationError
from .security import token_matches


LOGGER = logging.getLogger("ragsearch_service.http")


class LocalHTTPServer(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True


def create_http_server(service: SearchService) -> LocalHTTPServer:
    class Handler(BaseHTTPRequestHandler):
        server_version = "RAGSearchLocal/0.1"
        sys_version = ""

        def log_message(self, format_string: str, *args: object) -> None:
            LOGGER.info("%s - %s", self.client_address[0], format_string % args)

        def _write_json(self, status: int, payload: object) -> None:
            encoded = json.dumps(
                payload,
                ensure_ascii=False,
                separators=(",", ":"),
            ).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(encoded)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            self.wfile.write(encoded)

        def _authorized(self) -> bool:
            return token_matches(service.token, self.headers.get("X-RAGSearch-Token"))

        def _require_authorized(self) -> bool:
            if self._authorized():
                return True
            self._write_json(
                HTTPStatus.UNAUTHORIZED,
                {"error": "unauthorized", "message": "Missing or invalid X-RAGSearch-Token"},
            )
            return False

        def _read_json(self) -> Any:
            raw_length = self.headers.get("Content-Length")
            if raw_length is None:
                raise ValidationError("Content-Length is required")
            try:
                length = int(raw_length)
            except ValueError as exc:
                raise ValidationError("Invalid Content-Length") from exc
            if length <= 0:
                raise ValidationError("JSON request body is required")
            if length > service.settings.request_limit_bytes:
                raise OverflowError("Request body exceeds the configured limit")
            media_type = self.headers.get_content_type()
            if media_type != "application/json":
                raise ValidationError("Content-Type must be application/json")
            payload = self.rfile.read(length)
            try:
                return json.loads(payload.decode("utf-8"))
            except (UnicodeDecodeError, json.JSONDecodeError) as exc:
                raise ValidationError("Request body is not valid UTF-8 JSON") from exc

        def do_GET(self) -> None:
            path = urlsplit(self.path).path
            if path == "/health":
                try:
                    self._write_json(HTTPStatus.OK, service.health())
                except Exception:
                    LOGGER.exception("Health check failed")
                    self._write_json(
                        HTTPStatus.SERVICE_UNAVAILABLE,
                        {"status": "error", "protocol": PROTOCOL_VERSION},
                    )
                return
            if path == "/v1/stats":
                if not self._require_authorized():
                    return
                try:
                    self._write_json(HTTPStatus.OK, service.stats())
                except Exception:
                    LOGGER.exception("Stats request failed")
                    self._write_json(
                        HTTPStatus.INTERNAL_SERVER_ERROR,
                        {"error": "internal_error"},
                    )
                return
            self._write_json(HTTPStatus.NOT_FOUND, {"error": "not_found"})

        def do_POST(self) -> None:
            path = urlsplit(self.path).path
            if path not in {"/v1/documents", "/v1/search"}:
                self._write_json(HTTPStatus.NOT_FOUND, {"error": "not_found"})
                return
            if not self._require_authorized():
                return
            try:
                payload = self._read_json()
                if path == "/v1/documents":
                    response = service.ingest_document(payload)
                else:
                    response = service.search(payload)
                self._write_json(HTTPStatus.OK, response)
            except OverflowError as exc:
                self._write_json(
                    HTTPStatus.REQUEST_ENTITY_TOO_LARGE,
                    {"error": "request_too_large", "message": str(exc)},
                )
            except ValidationError as exc:
                self._write_json(
                    HTTPStatus.BAD_REQUEST,
                    {"error": "invalid_request", "message": str(exc)},
                )
            except Exception:
                LOGGER.exception("API request failed")
                self._write_json(
                    HTTPStatus.INTERNAL_SERVER_ERROR,
                    {"error": "internal_error"},
                )

        def do_DELETE(self) -> None:
            path = urlsplit(self.path).path
            if path != "/v1/index":
                self._write_json(HTTPStatus.NOT_FOUND, {"error": "not_found"})
                return
            if not self._require_authorized():
                return
            try:
                self._write_json(HTTPStatus.OK, service.clear_index())
            except Exception:
                LOGGER.exception("Index clear request failed")
                self._write_json(
                    HTTPStatus.INTERNAL_SERVER_ERROR,
                    {"error": "internal_error"},
                )

    return LocalHTTPServer((service.settings.host, service.settings.port), Handler)
