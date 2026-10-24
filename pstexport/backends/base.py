"""Backend contract. Both concrete backends implement exactly this surface."""
from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, Iterator

from ..models import MessageRecord


class BaseBackend(ABC):
    """A read-only PST source.

    Implementations must be *fail-soft*: `iter_messages` should skip a bad
    message rather than abort the folder, and `read_message` may raise so the
    engine can count the failure.
    """

    name: str = "base"
    description: str = ""

    @classmethod
    def available(cls) -> bool:
        """True when this backend's dependencies can be imported."""
        return False

    @abstractmethod
    def open(self, path: Path) -> None:
        """Open the store. Raises PSTOpenError on total failure."""

    @abstractmethod
    def close(self) -> None: ...

    @abstractmethod
    def root_folder(self) -> Any: ...

    @abstractmethod
    def folder_name(self, folder: Any) -> str: ...

    @abstractmethod
    def subfolders(self, folder: Any) -> Iterator[Any]: ...

    @abstractmethod
    def message_count(self, folder: Any) -> int: ...

    @abstractmethod
    def iter_messages(self, folder: Any) -> Iterator[Any]: ...

    @abstractmethod
    def read_message(self, message: Any, folder_path: str) -> MessageRecord: ...

    def health(self) -> dict[str, Any]:
        return {}
