"""Streaming CSV writer. Constant memory regardless of mailbox size."""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Sequence

from .errors import OutputError
from .models import CSV_COLUMNS, MessageRecord

_FORMULA_TRIGGERS = ("=", "+", "@", "\t")
_NUMERIC = ("-",)


def excel_safe(value: str) -> str:
    """Defuse CSV formula injection without mangling negative numbers."""
    if not value:
        return value
    first = value[0]
    if first in _FORMULA_TRIGGERS:
        return "'" + value
    if first in _NUMERIC:
        try:
            float(value)
            return value
        except ValueError:
            return "'" + value
    return value


class BaseWriter:
    def write(self, record: MessageRecord) -> None: ...
    def close(self) -> None: ...
    @property
    def bytes_written(self) -> int: return 0
    def __enter__(self): return self
    def __exit__(self, *exc): self.close()


class NullWriter(BaseWriter):
    """Used by --dry-run."""

    def __init__(self) -> None:
        self.count = 0

    def write(self, record: MessageRecord) -> None:
        self.count += 1


class CsvStreamWriter(BaseWriter):
    """Row-at-a-time CSV emitter.

    - `utf-8-sig` so Excel opens non-ASCII correctly on a double-click.
    - `QUOTE_ALL` so embedded delimiters, quotes and newlines are always safe.
    - CRLF line terminator for maximum Excel/Windows compatibility.
    """

    def __init__(
        self,
        path: Path,
        columns: Sequence[str] = CSV_COLUMNS,
        encoding: str = "utf-8-sig",
        delimiter: str = ",",
        excel_safe_cells: bool = True,
        flush_every: int = 250,
    ) -> None:
        self.path = Path(path)
        self._excel_safe = excel_safe_cells
        self._flush_every = max(1, flush_every)
        self._since_flush = 0
        self._bytes = 0

        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self._handle = self.path.open("w", newline="", encoding=encoding, errors="replace")
        except OSError as exc:
            raise OutputError(f"cannot open output file {self.path}: {exc}") from exc

        self._writer = csv.writer(
            self._handle,
            delimiter=delimiter,
            quoting=csv.QUOTE_ALL,
            lineterminator="\r\n",
        )
        self._write_row(list(columns))

    def _write_row(self, row: list) -> None:
        self._writer.writerow(row)
        # Cheap size accounting for the dashboard gauge.
        self._bytes += sum(len(str(cell).encode("utf-8", "replace")) + 3 for cell in row) + 2

    def write(self, record: MessageRecord) -> None:
        row = record.as_row()
        if self._excel_safe:
            row = [excel_safe(cell) if isinstance(cell, str) else cell for cell in row]
        self._write_row(row)

        self._since_flush += 1
        if self._since_flush >= self._flush_every:
            self._handle.flush()
            self._since_flush = 0

    @property
    def bytes_written(self) -> int:
        return self._bytes

    def close(self) -> None:
        handle = getattr(self, "_handle", None)
        if handle is None or handle.closed:
            return
        try:
            handle.flush()
            handle.close()
        except OSError as exc:
            raise OutputError(f"error finalising {self.path}: {exc}") from exc
