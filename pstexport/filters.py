"""Deduplication and spam exclusion. Both are pure, side-effect-free predicates."""
from __future__ import annotations

import hashlib
import re
from typing import Iterable, Optional

from .models import MessageRecord
from .utils import safe_text

# Matched against whole path *segments*, so "Junk E-mail" and "Junk" both hit
# while a folder like "JunkDrawer" does not.
SPAM_FOLDER_TOKENS: frozenset[str] = frozenset({
    "junk e-mail",
    "junk email",
    "junk-e-mail",
    "junk",
    "spam",
    "bulk mail",
    "unwanted email",
    "quarantine",
    "spam suspect",
})

_SPAM_HEADERS = ("x-spam-flag", "x-spam-status", "x-spam", "x-spam-flag-status")
_SCL_HEADERS = ("x-ms-exchange-organization-scl", "x-microsoft-antispam")
_SCORE_HEADERS = ("x-spam-score", "x-spamd-score", "x-spam-level")

_PATH_SPLIT = re.compile(r"[\\/]")


class Deduplicator:
    """Remembers a digest per message. Digest size is 16 bytes, so ~1M
    messages costs roughly 100 MB of `set` overhead — acceptable, and far
    cheaper than retaining bodies."""

    def __init__(self, mode: str = "message_id", max_keys: int = 3_000_000) -> None:
        if mode not in ("message_id", "content"):
            raise ValueError(f"unknown dedupe mode: {mode!r}")
        self.mode = mode
        self._seen: set[bytes] = set()
        self._max_keys = max_keys
        self._overflowed = False

    @property
    def size(self) -> int:
        return len(self._seen)

    @property
    def overflowed(self) -> bool:
        return self._overflowed

    def signature(self, record: MessageRecord) -> bytes:
        if self.mode == "message_id" and record.message_id:
            basis = "mid\x1f" + record.message_id.strip().strip("<>").lower()
        else:
            # Content mode, and also the fallback when Message-ID is absent.
            stamp = record.date.strftime("%Y-%m-%d %H:%M") if record.date else ""
            basis = "\x1f".join((
                "cnt",
                record.subject.strip().lower(),
                record.sender_email.strip().lower(),
                stamp,
                record.body_plain[:4096],
            ))
        return hashlib.blake2b(basis.encode("utf-8", "replace"), digest_size=16).digest()

    def is_duplicate(self, record: MessageRecord) -> bool:
        key = self.signature(record)
        if key in self._seen:
            return True
        if len(self._seen) >= self._max_keys:
            self._overflowed = True
        self._seen.add(key)
        return False


class SpamFilter:
    """Two independent signals: folder placement and transport headers."""

    def __init__(
        self,
        folder_tokens: Iterable[str] = SPAM_FOLDER_TOKENS,
        use_headers: bool = True,
    ) -> None:
        self.folder_tokens = frozenset(t.lower() for t in folder_tokens)
        self.use_headers = use_headers

    # -- folder signal ---------------------------------------------------
    def folder_is_spam(self, folder_path: str) -> bool:
        segments = [seg.strip().lower() for seg in _PATH_SPLIT.split(folder_path)]
        return any(seg in self.folder_tokens for seg in segments if seg)

    # -- header signal ---------------------------------------------------
    @staticmethod
    def _score(value: str) -> Optional[float]:
        match = re.search(r"[-+]?\d+(?:\.\d+)?", value)
        return float(match.group()) if match else None

    def headers_are_spam(self, headers_raw: str) -> bool:
        from .utils import parse_headers  # local import avoids a cycle at module load

        headers = parse_headers(headers_raw)

        for name in _SPAM_HEADERS:
            value = headers.get(name, "").strip().lower()
            if value in ("yes", "true", "1", "spam"):
                return True

        for name in _SCL_HEADERS:
            value = headers.get(name, "")
            score = self._score(value)
            if score is not None and score >= 5:
                return True

        for name in _SCORE_HEADERS:
            value = headers.get(name, "")
            score = self._score(value)
            if score is not None and score >= 5.0:
                return True

        return False

    def is_spam(self, record: MessageRecord, folder_path: str) -> bool:
        if self.folder_is_spam(folder_path):
            return True
        if self.use_headers and record.headers_raw:
            return self.headers_are_spam(record.headers_raw)
        return False
