from __future__ import annotations

import html.parser
import zipfile
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree


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
class ExtractedPart:
    text: str
    status: str
    error: str | None


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


def _extract_docx(payload: bytes, limit_bytes: int) -> str:
    with zipfile.ZipFile(BytesIO(payload)) as archive:
        try:
            info = archive.getinfo("word/document.xml")
        except KeyError as exc:
            raise ValueError("DOCX has no word/document.xml") from exc
        if info.file_size > limit_bytes:
            raise OverflowError("DOCX document XML exceeds the extraction limit")
        document_xml = archive.read(info)

    root = ElementTree.fromstring(document_xml)
    parts: list[str] = []
    for element in root.iter():
        local_name = element.tag.rsplit("}", 1)[-1]
        if local_name == "t" and element.text:
            parts.append(element.text)
        elif local_name in {"p", "br", "tab"}:
            parts.append("\n" if local_name != "tab" else "\t")
    return " ".join(parts)


def extract_part(
    content: bytes | None,
    *,
    name: str,
    media_type: str,
    limit_bytes: int,
    text_limit_chars: int,
) -> ExtractedPart:
    """Extract searchable text from inline part bytes.

    The HTTP boundary owns decoding base64.  This function deliberately accepts
    bytes rather than filesystem paths, so the service has no connector spool or
    shared-filesystem contract.
    """

    if content is None:
        return ExtractedPart("", "not_provided", None)
    if len(content) > limit_bytes:
        return ExtractedPart(
            "",
            "too_large",
            f"Part content exceeds extraction limit ({limit_bytes} bytes)",
        )

    suffix = Path(name).suffix.casefold()
    folded_media_type = media_type.casefold()
    supported = (
        suffix in _TEXT_EXTENSIONS
        or suffix == ".docx"
        or folded_media_type.startswith("text/")
        or folded_media_type in {"application/json", "application/xml"}
    )
    if not supported:
        return ExtractedPart("", "unsupported", None)

    try:
        if suffix == ".docx":
            text = _extract_docx(content, limit_bytes)
        else:
            text = _decode_bytes(content)
            if suffix in {".html", ".htm"} or folded_media_type == "text/html":
                parser = _HTMLTextExtractor()
                parser.feed(text)
                text = parser.text()
        if len(text) > text_limit_chars:
            return ExtractedPart(
                text[:text_limit_chars],
                "extracted_truncated",
                f"Extracted text truncated at {text_limit_chars} characters",
            )
        return ExtractedPart(text, "extracted", None)
    except OverflowError as exc:
        return ExtractedPart("", "too_large", str(exc))
    except (ValueError, zipfile.BadZipFile, ElementTree.ParseError) as exc:
        return ExtractedPart("", "error", str(exc))
