from datetime import datetime
from pathlib import Path
import csv

from pstexport.filters import Deduplicator, SpamFilter
from pstexport.models import MessageRecord
from pstexport.writers import CsvStreamWriter, excel_safe


def test_formula_injection_is_escaped_but_negative_numbers_are_not():
    assert excel_safe("=1+1").startswith("'")
    assert excel_safe("@SUM(A1:A2)").startswith("'")
    assert excel_safe("-12.5") == "-12.5"
    assert excel_safe("-cmd") == "'-cmd"


def test_spam_filter_matches_folder_segments_not_substrings():
    spam_filter = SpamFilter()
    assert spam_filter.folder_is_spam("Inbox/Junk E-mail")
    assert not spam_filter.folder_is_spam("Inbox/JunkDrawer")


def test_deduplicator_uses_message_id_case_insensitively():
    deduplicator = Deduplicator("message_id")
    first = MessageRecord(message_id="ABC-123")
    duplicate = MessageRecord(message_id="abc-123")
    assert deduplicator.is_duplicate(first) is False
    assert deduplicator.is_duplicate(duplicate) is True


def test_csv_writer_emits_header_and_message(tmp_path: Path):
    output = tmp_path / "mail.csv"
    record = MessageRecord(
        date=datetime(2026, 10, 24, 12, 0, 0),
        sender_name="Test Sender",
        sender_email="sender@example.com",
        subject="Hello",
        body_plain="First line\nSecond line",
        message_id="test-id",
    )
    with CsvStreamWriter(output) as writer:
        writer.write(record)
    with output.open("r", encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))
    assert len(rows) == 1
    assert rows[0]["subject"] == "Hello"
    assert rows[0]["body_plain"] == "First line\nSecond line"
    assert rows[0]["message_id"] == "test-id"
