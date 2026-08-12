from __future__ import annotations

import datetime as dt
import re


_ISO_8601_WITH_OFFSET = re.compile(
    r"\A\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}"
    r"(?:\.\d{1,6})?(?:Z|[+-]\d{2}:\d{2})\Z"
)


def canonical_utc_timestamp(value: str) -> str:
    """Validate the public timestamp grammar and return fixed-width UTC."""
    if not _ISO_8601_WITH_OFFSET.fullmatch(value):
        raise ValueError("expected an ISO-8601 timestamp with an explicit UTC offset")

    parseable = value[:-1] + "+00:00" if value.endswith("Z") else value
    try:
        parsed = dt.datetime.fromisoformat(parseable)
    except ValueError as exc:
        raise ValueError(
            "expected an ISO-8601 timestamp with an explicit UTC offset"
        ) from exc
    if parsed.utcoffset() is None:
        raise ValueError("expected an ISO-8601 timestamp with an explicit UTC offset")

    return (
        parsed.astimezone(dt.timezone.utc)
        .isoformat(timespec="microseconds")
        .replace("+00:00", "Z")
    )
