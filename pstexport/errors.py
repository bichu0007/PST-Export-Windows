"""Typed exception hierarchy so callers can react precisely."""
from __future__ import annotations


class PSTExportError(Exception):
    """Base class for every toolkit error."""


class BackendUnavailableError(PSTExportError):
    """No usable PST reader is installed."""


class PSTOpenError(PSTExportError):
    """The PST could not be opened at all (fatal)."""


class PSTReadError(PSTExportError):
    """A folder or message failed to read (non-fatal)."""


class OutputError(PSTExportError):
    """Problem writing the destination file."""


class UserAbort(PSTExportError):
    """The operator cancelled the wizard."""
