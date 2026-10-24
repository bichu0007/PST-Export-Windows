"""libpff (`pypff`) backend — fast, native, no Outlook required."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Iterator, Optional

from ..errors import PSTOpenError
from ..models import MessageRecord
from ..utils import (
    clean_body,
    clean_field,
    decode_mime,
    html_to_text,
    parse_headers,
    safe_text,
    split_addresses,
)
from .base import BaseBackend

_EPOCH_FLOOR = datetime(1979, 1, 1)


def _call(obj: Any, method: str, default: Any = None) -> Any:
    """Invoke a pypff getter defensively — the binding raises liberally."""
    func = getattr(obj, method, None)
    if func is None:
        return default
    try:
        return func()
    except Exception:
        return default


class PypffBackend(BaseBackend):
    name = "pypff"
    description = "libpff / pypff native reader"

    def __init__(self) -> None:
        self._module: Any = None
        self._store: Any = None

    @classmethod
    def available(cls) -> bool:
        try:
            import pypff  # noqa: F401
            return True
        except Exception:
            return False

    # -- lifecycle -------------------------------------------------------
    def open(self, path: Path) -> None:
        import pypff

        self._module = pypff
        store = pypff.file()
        try:
            store.open(str(path))
        except Exception as exc:
            raise PSTOpenError(
                f"pypff could not open {path.name}: {exc}. "
                "If the file is corrupt, run SCANPST.EXE, or retry with --backend outlook."
            ) from exc
        self._store = store

    def close(self) -> None:
        if self._store is not None:
            try:
                self._store.close()
            except Exception:
                pass
            self._store = None

    # -- structure -------------------------------------------------------
    def root_folder(self) -> Any:
        return self._store.get_root_folder()

    def folder_name(self, folder: Any) -> str:
        name = _call(folder, "get_name", None)
        text = clean_field(name)
        return text or "(unnamed)"

    def subfolders(self, folder: Any) -> Iterator[Any]:
        count = _call(folder, "get_number_of_sub_folders", 0) or 0
        for index in range(count):
            try:
                yield folder.get_sub_folder(index)
            except Exception:
                continue

    def message_count(self, folder: Any) -> int:
        return int(_call(folder, "get_number_of_sub_messages", 0) or 0)

    def iter_messages(self, folder: Any) -> Iterator[Any]:
        count = self.message_count(folder)
        for index in range(count):
            try:
                yield folder.get_sub_message(index)
            except Exception:
                # Bad directory entry — surface as None so the engine counts it.
                yield None

    # -- extraction ------------------------------------------------------
    def _pick_datetime(self, message: Any) -> Optional[datetime]:
        for getter in ("get_delivery_time", "get_client_submit_time", "get_modification_time"):
            value = _call(message, getter, None)
            if isinstance(value, datetime) and value >= _EPOCH_FLOOR:
                return value
        return None

    def read_message(self, message: Any, folder_path: str) -> MessageRecord:
        if message is None:
            raise ValueError("null message handle")

        record = MessageRecord(folder_path=folder_path)

        # Transport headers give us the trustworthy RFC 5322 view.
        headers_raw = safe_text(_call(message, "get_transport_headers", ""))
        if not headers_raw:
            headers_raw = safe_text(_call(message, "get_headers", ""))
        record.headers_raw = headers_raw
        headers = parse_headers(headers_raw)

        # --- subject ---------------------------------------------------
        subject = _call(message, "get_subject", "") or headers.get("subject", "")
        record.subject = decode_mime(clean_field(subject, limit=1000))

        # --- sender ----------------------------------------------------
        sender_name = _call(message, "get_sender_name", "") or ""
        from_header = decode_mime(headers.get("from", ""))
        if not sender_name and from_header:
            sender_name = from_header.split("<")[0]
        record.sender_name = clean_field(sender_name, limit=320)

        sender_email = headers.get("from", "")
        if "<" in sender_email and ">" in sender_email:
            sender_email = sender_email[sender_email.rfind("<") + 1: sender_email.rfind(">")]
        elif sender_name and sender_email.strip() == sender_name.strip():
            sender_email = ""
        record.sender_email = clean_field(sender_email, limit=320)

        # --- recipients ------------------------------------------------
        to_list = headers.get("to", "")
        cc_list = headers.get("cc", "")

        if not to_list:
            # Fall back to the MAPI recipient table when headers are absent.
            collected: list[str] = []
            count = int(_call(message, "get_number_of_recipients", 0) or 0)
            for index in range(count):
                try:
                    recipient = message.get_recipient(index)
                except Exception:
                    continue
                address = clean_field(_call(recipient, "get_email_address", ""))
                if address:
                    collected.append(address)
            to_list = "; ".join(collected)

        record.recipients_to = split_addresses(decode_mime(to_list), limit=30)
        record.recipients_cc = split_addresses(decode_mime(cc_list), limit=30)

        # --- date ------------------------------------------------------
        record.date = self._pick_datetime(message)
        if record.date is None:
            record.date = _parse_header_date(headers.get("date", ""))

        # --- body ------------------------------------------------------
        body = safe_text(_call(message, "get_plain_text_body", ""))
        if not body.strip():
            html = safe_text(_call(message, "get_html_body", ""))
            if html.strip():
                body = html_to_text(html)
        if not body.strip():
            rtf = safe_text(_call(message, "get_rtf_body", ""))
            if rtf.strip():
                body = clean_body(rtf)
        record.body_plain = clean_body(body)

        # --- attachments -----------------------------------------------
        attachment_count = int(_call(message, "get_number_of_attachments", 0) or 0)
        record.attachment_count = max(0, attachment_count)
        record.has_attachments = record.attachment_count > 0

        # --- identity / size -------------------------------------------
        record.message_id = clean_field(
            headers.get("message-id", "") or _call(message, "get_identifier", "") or "",
            limit=400,
        ).strip("<>")
        record.size_bytes = int(_call(message, "get_size", 0) or 0)

        return record


def _parse_header_date(value: str) -> Optional[datetime]:
    if not value:
        return None
    try:
        from email.utils import parsedate_to_datetime

        parsed = parsedate_to_datetime(value)
        if parsed is not None:
            return parsed.replace(tzinfo=None) if parsed.tzinfo else parsed
    except Exception:
        pass
    return None
