"""Backend registry and resolver."""
from __future__ import annotations

from ..errors import BackendUnavailableError
from .base import BaseBackend


def _registry() -> dict[str, type[BaseBackend]]:
    from .outlook_backend import OutlookBackend
    from .pypff_backend import PypffBackend

    return {"pypff": PypffBackend, "outlook": OutlookBackend}


def available_backends() -> dict[str, type[BaseBackend]]:
    found: dict[str, type[BaseBackend]] = {}
    for name, cls in _registry().items():
        try:
            if cls.available():
                found[name] = cls
        except Exception:
            continue
    return found


def load_backend(name: str = "auto") -> BaseBackend:
    available = available_backends()

    if name == "auto":
        for preferred in ("pypff", "outlook"):
            if preferred in available:
                return available[preferred]()
        raise BackendUnavailableError(
            "No PST reader available.\n"
            "  Install one of:\n"
            "    pip install libpff-python   (recommended)\n"
            "    pip install pywin32         (requires Outlook installed)"
        )

    if name not in available:
        raise BackendUnavailableError(
            f"Backend '{name}' is not usable. Available: "
            f"{', '.join(sorted(available)) or 'none'}"
        )
    return available[name]()


__all__ = ["BaseBackend", "available_backends", "load_backend"]
