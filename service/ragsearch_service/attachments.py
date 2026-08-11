from __future__ import annotations

import html.parser
import os
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree

from .errors import ValidationError


_TEXT_EXTENSIONS = {
    ".csv",
    ".htm",
    ".html",
    ".json",
    ".log",
    ".md",
    ".pub",
    ".rst",
    ".text",
    ".txt",
    ".xml",
    ".yaml",
    ".yml",
}


@dataclass(frozen=True)
class ExtractedAttachment:
    text: str
    status: str
    error: str | None
    safe_path: Path | None


class _HTMLTextExtractor(html.parser.HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        self.parts.append(data)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.casefold() in {"br", "p", "div", "li", "tr", "h1", "h2", "h3"}:
            self.parts.append("\n")

    def text(self) -> str:
        return " ".join(self.parts)


def _safe_spool_path(raw_path: str, spool_dir: Path) -> Path:
    try:
        root = spool_dir.resolve(strict=True)
        candidate = Path(raw_path).resolve(strict=True)
    except (OSError, RuntimeError) as exc:
        raise ValidationError(f"Attachment temp_path does not exist: {raw_path}") from exc

    try:
        common = os.path.commonpath((os.path.normcase(str(root)), os.path.normcase(str(candidate))))
    except ValueError as exc:
        raise ValidationError("Attachment temp_path is outside the configured spool directory") from exc
    if common != os.path.normcase(str(root)) or not candidate.is_file():
        raise ValidationError("Attachment temp_path is outside the configured spool directory")
    return candidate


def _decode_bytes(payload: bytes) -> str:
    if payload.startswith((b"\xff\xfe", b"\xfe\xff")):
        return payload.decode("utf-16")
    if payload.startswith(b"\xef\xbb\xbf"):
        return payload.decode("utf-8-sig")
    for encoding in ("utf-8", "cp1251", "latin-1"):
        try:
            return payload.decode(encoding)
        except UnicodeDecodeError:
            continue
    return payload.decode("utf-8", errors="replace")


def _extract_docx(path: Path, limit_bytes: int) -> str:
    with zipfile.ZipFile(path) as archive:
        try:
            info = archive.getinfo("word/document.xml")
        except KeyError as exc:
            raise ValueError("DOCX has no word/document.xml") from exc
        if info.file_size > limit_bytes:
            raise OverflowError("DOCX document XML exceeds the extraction limit")
        payload = archive.read(info)

    root = ElementTree.fromstring(payload)
    parts: list[str] = []
    for element in root.iter():
        local_name = element.tag.rsplit("}", 1)[-1]
        if local_name == "t" and element.text:
            parts.append(element.text)
        elif local_name in {"p", "br", "tab"}:
            parts.append("\n" if local_name != "tab" else "\t")
    return " ".join(parts)


def extract_attachment(
    attachment: dict[str, object],
    *,
    spool_dir: Path,
    limit_bytes: int,
) -> ExtractedAttachment:
    raw_path = attachment.get("temp_path")
    if raw_path in {None, ""}:
        return ExtractedAttachment("", "not_provided", None, None)
    if not isinstance(raw_path, str):
        raise ValidationError("Attachment temp_path must be a string or null")

    safe_path = _safe_spool_path(raw_path, spool_dir)
    declared_name = attachment.get("name")
    declared_suffix = (
        Path(declared_name).suffix.casefold() if isinstance(declared_name, str) else ""
    )
    suffix = declared_suffix or safe_path.suffix.casefold()
    content_type = str(attachment.get("content_type") or "").casefold()
    supported = (
        suffix in _TEXT_EXTENSIONS
        or suffix == ".docx"
        or content_type.startswith("text/")
        or content_type in {"application/json", "application/xml"}
    )
    if not supported:
        return ExtractedAttachment("", "unsupported", None, safe_path)

    try:
        size = safe_path.stat().st_size
        if size > limit_bytes:
            return ExtractedAttachment(
                "",
                "too_large",
                f"Attachment exceeds extraction limit ({limit_bytes} bytes)",
                safe_path,
            )
        if suffix == ".docx":
            text = _extract_docx(safe_path, limit_bytes)
        else:
            text = _decode_bytes(safe_path.read_bytes())
            if suffix in {".html", ".htm"} or content_type == "text/html":
                parser = _HTMLTextExtractor()
                parser.feed(text)
                text = parser.text()
        return ExtractedAttachment(text, "extracted", None, safe_path)
    except OverflowError as exc:
        return ExtractedAttachment("", "too_large", str(exc), safe_path)
    except (OSError, ValueError, zipfile.BadZipFile, ElementTree.ParseError) as exc:
        return ExtractedAttachment("", "error", str(exc), safe_path)
