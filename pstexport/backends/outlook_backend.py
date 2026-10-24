"""Outlook COM (`pywin32`) backend — slow but extremely tolerant of odd stores."""
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

# MAPI property DASL tags
PR_TRANSPORT_HEADERS = "http://schemas.microsoft.com/mapi/proptags/0x007D001E"
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptags/0x1035001E"

OL_MAIL_CLASS = 43
OL_STORE_UNICODE = 2
RECIPIENT_TYPE = {1: "to", 2: "cc", 3: "bcc"}


class OutlookBackend(BaseBackend):
    name = "outlook"
    description = "Microsoft Outlook COM automation"

    def __init__(self) -> None:
        self._pythoncom: Any = None
        self._namespace: Any = None
        self._store: Any = None
        self._root: Any = None
        self._com_initialised = False

    @classmethod
    def available(cls) -> bool:
        try:
            import win32com.client  # noqa: F401
            import pythoncom  # noqa: F401
            return True
        except Exception:
            return False

    # -- lifecycle -------------------------------------------------------
    def open(self, path: Path) -> None:
        import pythoncom
        import win32com.client

        self._pythoncom = pythoncom
        pythoncom.CoInitialize()
        self._com_initialised = True

        try:
            app = win32com.client.Dispatch("Outlook.Application")
            namespace = app.GetNamespace("MAPI")
        except Exception as exc:
            raise PSTOpenError(
                f"Cannot attach to Outlook via COM: {exc}. "
                "Is Outlook installed and licensed?"
            ) from exc

        # AddStoreEx(..., 2) mounts the store read-write; Outlook may still
        # have it open, which is fine.
        try:
            namespace.AddStoreEx(str(path), OL_STORE_UNICODE)
        except Exception:
            try:
                namespace.AddStore(str(path))
            except Exception as exc:
                raise PSTOpenError(
                    f"Outlook refused to mount {path.name}: {exc}"
                ) from exc

        target = str(path).lower()
        store = None
        for candidate in namespace.Stores:
            try:
                if str(candidate.FilePath).lower() == target:
                    store = candidate
                    break
            except Exception:
                continue

        if store is None:
            raise PSTOpenError(
                f"Store mounted but could not be located in the profile: {path.name}"
            )

        self._namespace = namespace
        self._store = store
        self._root = store.GetRootFolder()

    def close(self) -> None:
        if self._namespace is not None and self._root is not None:
            try:
                self._namespace.RemoveStore(self._root)
            except Exception:
                pass
        if self._com_initialised:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_initialised = False
        self._namespace = self._store = self._root = None

    # -- structure -------------------------------------------------------
    def root_folder(self) -> Any:
        return self._root

    def folder_name(self, folder: Any) -> str:
        return clean_field(getattr(folder, "Name", "")) or "(unnamed)"

    def subfolders(self, folder: Any) -> Iterator[Any]:
        try:
            collection = folder.Folders
            count = collection.Count
        except Exception:
            return
        for index in range(1, count + 1):
            try:
                yield collection.Item(index)
            except Exception:
                continue

    def _mail_items(self, folder: Any) -> Any:
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]")
            return items
        except Exception:
            return folder.Items

    def message_count(self, folder: Any) -> int:
        try:
            return int(folder.Items.Count)
        except Exception:
            return 0

    def iter_messages(self, folder: Any) -> Iterator[Any]:
        try:
            items = self._mail_items(folder)
        except Exception:
            return
        for item in items:
            try:
                if getattr(item, "Class", None) != OL_MAIL_CLASS:
                    continue
            except Exception:
                continue
            yield item

    # -- extraction ------------------------------------------------------
    @staticmethod
    def _prop(item: Any, dasl: str) -> str:
        try:
            return safe_text(item.PropertyAccessor.GetProperty(dasl))
        except Exception:
            return ""

    def read_message(self, item: Any, folder_path: str) -> MessageRecord:
        record = MessageRecord(folder_path=folder_path)

        headers_raw = self._prop(item, PR_TRANSPORT_HEADERS)
        record.headers_raw = headers_raw
        headers = parse_headers(headers_raw)

        record.subject = decode_mime(clean_field(getattr(item, "Subject", ""), limit=1000))

        sender_name = clean_field(getattr(item, "SenderName", ""), limit=320)
        record.sender_name = sender_name or clean_field(headers.get("from", "").split("<")[0], limit=320)

        address = clean_field(getattr(item, "SenderEmailAddress", ""), limit=320)
        if not address or not address.startswith("http"):
            record.sender_email = address or clean_field(headers.get("from", ""), limit=320)
        else:
            # Exchange X500 DN — prefer the SMTP header when available.
            record.sender_email = clean_field(headers.get("from", ""), limit=320)

        # Recipients via the MAPI recipient table (more reliable than headers).
        buckets: dict[str, list[str]] = {"to": [], "cc": [], "bcc": []}
        try:
            recipients = item.Recipients
            for index in range(1, recipients.Count + 1):
                recipient = recipients.Item(index)
                kind = RECIPIENT_TYPE.get(int(getattr(recipient, "Type", 1)), "to")
                addr = clean_field(getattr(recipient, "Address", ""))
                if addr:
                    buckets[kind].append(addr)
        except Exception:
            pass

        if not buckets["to"] and headers.get("to"):
            buckets["to"] = [headers["to"]]
        if not buckets["cc"] and headers.get("cc"):
            buckets["cc"] = [headers["cc"]]

        record.recipients_to = split_addresses("; ".join(buckets["to"]), limit=30)
        record.recipients_cc = split_addresses("; ".join(buckets["cc"]), limit=30)

        received = getattr(item, "ReceivedTime", None)
        if isinstance(received, datetime):
            record.date = received.replace(tzinfo=None)

        body = safe_text(getattr(item, "Body", ""))
        if not body.strip():
            html = safe_text(getattr(item, "HTMLBody", ""))
            if html.strip():
                body = html_to_text(html)
        record.body_plain = clean_body(body)

        try:
            count = int(item.Attachments.Count)
        except Exception:
            count = 0
        record.attachment_count = max(0, count)
        record.has_attachments = record.attachment_count > 0

        record.message_id = clean_field(
            self._prop(item, PR_INTERNET_MESSAGE_ID) or headers.get("message-id", ""),
            limit=400,
        ).strip("<>")

        try:
            record.size_bytes = int(item.Size)
        except Exception:
            record.size_bytes = 0

        return record
