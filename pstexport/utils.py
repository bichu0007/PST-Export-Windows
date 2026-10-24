"""Encoding coercion, header parsing, HTML flattening, formatting."""
from __future__ import annotations

import hashlib
import re
from html.parser import HTMLParser
from typing import Any, Optional

_NUL = re.compile(r"\x00")
_HSPACE = re.compile(r"[ \t\u00a0\u2007\u202f]+")
_MULTINL = re.compile(r"\n{3,}")
_HEADER_LINE = re.compile(r"^(?P<key>[A-Za-z0-9][A-Za-z0-9\-_.]*)\s*:\s*(?P<val>.*)$")

_BLOCK_TAGS = {
    "p", "div", "br", "tr", "li", "ul", "ol", "table", "h1", "h2", "h3",
    "h4", "h5", "h6", "blockquote", "pre", "hr", "section", "article",
}


# --------------------------------------------------------------------------
# Text coercion
# --------------------------------------------------------------------------
def safe_text(value: Any) -> str:
    """Coerce anything a backend might hand us into a clean `str`."""
    if value is None:
        return ""
    if isinstance(value, (bytes, bytearray, memoryview)):
        raw = bytes(value)
        for encoding in ("utf-8", "cp1252", "latin-1"):
            try:
                value = raw.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            value = raw.decode("utf-8", errors="replace")
    elif not isinstance(value, str):
        value = str(value)

    value = _NUL.sub("", value)
    return value.replace("\r\n", "\n").replace("\r", "\n")


def clean_field(value: Any, limit: int = 0) -> str:
    """Single-line field: collapse all whitespace, strip, optionally truncate."""
    text = _HSPACE.sub(" ", safe_text(value)).strip()
    if limit and len(text) > limit:
        text = text[:limit]
    return text


def clean_body(value: Any, limit: int = 0) -> str:
    """Multi-line body: preserve newlines, normalise horizontal whitespace."""
    text = _HSPACE.sub(" ", safe_text(value))
    text = _MULTINL.sub("\n\n", text)
    text = "\n".join(line.rstrip() for line in text.split("\n")).strip()
    if limit and len(text) > limit:
        text = text[:limit].rstrip() + " [...truncated]"
    return text


# --------------------------------------------------------------------------
# Headers
# --------------------------------------------------------------------------
def parse_headers(raw: Any) -> dict[str, str]:
    """Parse an RFC 5322 header block. Handles folded continuation lines.

    First occurrence of a repeated header wins, which is what you want for
    `Message-ID` and `From` in the presence of injected duplicates.
    """
    headers: dict[str, str] = {}
    current: Optional[str] = None

    for line in safe_text(raw).split("\n"):
        if not line.strip():
            break  # blank line terminates the header block
        if line[0] in " \t" and current is not None:
            headers[current] = headers[current] + " " + line.strip()
            continue
        match = _HEADER_LINE.match(line)
        if match:
            current = match.group("key").lower()
            headers.setdefault(current, match.group("val").strip())

    return headers


def decode_mime(value: str) -> str:
    """Decode RFC 2047 encoded-words (`=?utf-8?B?...?=`)."""
    if not value or "=?" not in value:
        return value
    try:
        from email.header import decode_header, make_header

        return str(make_header(decode_header(value)))
    except Exception:
        return value


def split_addresses(value: str, limit: int = 25) -> str:
    """Normalise an address list to `; `-separated, capped for CSV sanity."""
    if not value:
        return ""
    parts = [p.strip() for p in re.split(r"[;,]", value) if p.strip()]
    if len(parts) > limit:
        parts = parts[:limit] + [f"...(+{len(parts) - limit} more)"]
    return "; ".join(parts)


# --------------------------------------------------------------------------
# HTML
# --------------------------------------------------------------------------
class _HTMLFlattener(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._chunks: list[str] = []
        self._skip_depth = 0

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag in ("script", "style"):
            self._skip_depth += 1
        elif tag in _BLOCK_TAGS:
            self._chunks.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in ("script", "style"):
            self._skip_depth = max(0, self._skip_depth - 1)
        elif tag in _BLOCK_TAGS:
            self._chunks.append("\n")

    def handle_data(self, data: str) -> None:
        if not self._skip_depth:
            self._chunks.append(data)

    def result(self) -> str:
        return "".join(self._chunks)


def html_to_text(html: str) -> str:
    """Best-effort HTML → plain text without third-party parsers."""
    if not html:
        return ""
    parser = _HTMLFlattener()
    try:
        parser.feed(html)
        parser.close()
    except Exception:
        # Malformed markup: strip tags the crude way.
        return clean_body(re.sub(r"<[^>]+>", " ", html))
    return clean_body(parser.result())


# --------------------------------------------------------------------------
# Formatting
# --------------------------------------------------------------------------
def short_hash(text: str, size: int = 12) -> str:
    return hashlib.blake2b(text.encode("utf-8", "replace"), digest_size=size).hexdigest()


def human_size(num_bytes: float) -> str:
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if abs(num_bytes) < 1024.0:
            return f"{num_bytes:,.0f} {unit}" if unit == "B" else f"{num_bytes:,.1f} {unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:,.1f} PB"


def human_duration(seconds: Optional[float]) -> str:
    if seconds is None:
        return "--:--"
    seconds = int(max(0, seconds))
    hours, rem = divmod(seconds, 3600)
    minutes, secs = divmod(rem, 60)
    if hours:
        return f"{hours:d}:{minutes:02d}:{secs:02d}"
    return f"{minutes:02d}:{secs:02d}"


def truncate_middle(text: str, width: int) -> str:
    if width <= 4 or len(text) <= width:
        return text
    keep = width - 3
    left = keep // 2
    right = keep - left
    return f"{text[:left]}...{text[-right:]}"
