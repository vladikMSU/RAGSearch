from __future__ import annotations

import re
from collections.abc import Iterator


_WHITESPACE = re.compile(r"[ \t\f\v]+")


def normalize_text(value: str | None) -> str:
    if not value:
        return ""
    value = value.replace("\x00", " ").replace("\r\n", "\n").replace("\r", "\n")
    lines = (_WHITESPACE.sub(" ", line).strip() for line in value.split("\n"))
    return "\n".join(lines).strip()


def chunk_text(
    value: str | None,
    *,
    max_chars: int = 1_200,
    overlap_chars: int = 180,
) -> Iterator[str]:
    """Yield deterministic, bounded chunks without combining unrelated items."""
    if max_chars < 64:
        raise ValueError("max_chars must be at least 64")
    if overlap_chars < 0 or overlap_chars >= max_chars // 2:
        raise ValueError("overlap_chars must be non-negative and less than half max_chars")

    text = normalize_text(value)
    length = len(text)
    start = 0
    while start < length:
        hard_end = min(length, start + max_chars)
        end = hard_end
        if hard_end < length:
            minimum_break = start + max_chars // 2
            candidates = (
                text.rfind("\n", minimum_break, hard_end),
                text.rfind(". ", minimum_break, hard_end),
                text.rfind(" ", minimum_break, hard_end),
            )
            boundary = max(candidates)
            if boundary >= minimum_break:
                end = boundary + 1

        piece = text[start:end].strip()
        if piece:
            yield piece
        if end >= length:
            break
        start = max(start + 1, end - overlap_chars)

