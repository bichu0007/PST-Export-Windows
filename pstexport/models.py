"""Plain data containers shared across the toolkit."""
from __future__ import annotations

import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

# Column order for the emitted CSV. Extras beyond the required six fields
# are appended so the core six stay in the leftmost, easiest-to-read columns.
CSV_COLUMNS: tuple[str, ...] = (
    "date",
    "sender_name",
    "sender_email",
    "recipients_to",
    "recipients_cc",
    "subject",
    "body_plain",
    "has_attachments",
    "attachment_count",
    "folder_path",
    "message_id",
    "size_bytes",
)

DATE_FORMAT = "%Y-%m-%d %H:%M:%S"


@dataclass(slots=True)
class MessageRecord:
    """One normalised email. `headers_raw` never reaches the CSV."""

    date: Optional[datetime] = None
    sender_name: str = ""
    sender_email: str = ""
    recipients_to: str = ""
    recipients_cc: str = ""
    subject: str = ""
    body_plain: str = ""
    has_attachments: bool = False
    attachment_count: int = 0
    folder_path: str = ""
    message_id: str = ""
    size_bytes: int = 0
    headers_raw: str = ""

    def as_row(self) -> list[str | int]:
        return [
            self.date.strftime(DATE_FORMAT) if self.date else "",
            self.sender_name,
            self.sender_email,
            self.recipients_to,
            self.recipients_cc,
            self.subject,
            self.body_plain,
            "yes" if self.has_attachments else "no",
            self.attachment_count,
            self.folder_path,
            self.message_id,
            self.size_bytes,
        ]


@dataclass(slots=True)
class RunConfig:
    """Everything the engine needs. Produced by the wizard or by CLI flags."""

    pst_path: Path
    output_path: Path
    remove_duplicates: bool = True
    dedupe_mode: str = "message_id"          # "message_id" | "content"
    filter_spam: bool = True
    include_subfolders: bool = True
    exclude_folders: list[str] = field(default_factory=list)
    backend: str = "auto"                    # "auto" | "pypff" | "outlook"
    max_body_chars: int = 0                  # 0 = unlimited
    excel_safe: bool = True
    dry_run: bool = False
    limit: int = 0                           # 0 = no cap on exported rows
    log_path: Optional[Path] = None
    csv_encoding: str = "utf-8-sig"
    csv_delimiter: str = ","


@dataclass
class Stats:
    """Live counters. Mutated by the worker thread, read by the render thread."""

    phase: str = "idle"
    backend: str = ""
    pst_name: str = ""

    folders_discovered: int = 0
    folders_total: int = 0
    folders_scanned: int = 0

    messages_total: int = 0
    messages_seen: int = 0

    exported: int = 0
    duplicates: int = 0
    spam: int = 0
    errors: int = 0
    unreadable_folders: int = 0

    bytes_written: int = 0
    current_folder: str = ""

    started_at: float = field(default_factory=time.monotonic)
    finished_at: Optional[float] = None

    # ---- derived -------------------------------------------------------
    @property
    def elapsed(self) -> float:
        end = self.finished_at if self.finished_at is not None else time.monotonic()
        return max(0.0, end - self.started_at)

    @property
    def rate(self) -> float:
        """Messages processed per second."""
        elapsed = self.elapsed
        return (self.messages_seen / elapsed) if elapsed > 0.05 else 0.0

    @property
    def eta(self) -> Optional[float]:
        rate = self.rate
        if rate <= 0 or self.messages_total <= 0:
            return None
        remaining = max(0, self.messages_total - self.messages_seen)
        return remaining / rate

    @property
    def completion(self) -> float:
        if self.messages_total > 0:
            return min(1.0, self.messages_seen / self.messages_total)
        if self.folders_total > 0:
            return min(1.0, self.folders_scanned / self.folders_total)
        return 0.0
