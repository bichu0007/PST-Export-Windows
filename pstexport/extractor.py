"""Two-phase extraction engine: index the tree, then stream records out."""
from __future__ import annotations

import threading
import time
from pathlib import Path
from typing import Iterator, Optional

from .backends.base import BaseBackend
from .errors import PSTReadError
from .filters import Deduplicator, SpamFilter
from .models import MessageRecord, RunConfig, Stats
from .utils import clean_body, clean_field
from .writers import BaseWriter


class ExtractionEngine:
    """Traversal + filtering. Never raises for a single bad folder or message."""

    def __init__(
        self,
        config: RunConfig,
        backend: BaseBackend,
        writer: BaseWriter,
        stats: Stats,
        dashboard=None,
    ) -> None:
        self.config = config
        self.backend = backend
        self.writer = writer
        self.stats = stats
        self.dash = dashboard

        self._stop = threading.Event()
        self._lock = threading.Lock()

        self.deduplicator: Optional[Deduplicator] = (
            Deduplicator(mode=config.dedupe_mode) if config.remove_duplicates else None
        )
        self.spam_filter: Optional[SpamFilter] = (
            SpamFilter() if config.filter_spam else None
        )

        self._repair_log: Optional[Path] = config.log_path
        self._repair_handle = None
        self._excluded = {name.strip().lower() for name in config.exclude_folders if name.strip()}

    # ------------------------------------------------------------------
    # Control
    # ------------------------------------------------------------------
    def request_stop(self) -> None:
        self._stop.set()

    @property
    def cancelled(self) -> bool:
        return self._stop.is_set()

    # ------------------------------------------------------------------
    # Logging
    # ------------------------------------------------------------------
    def _log(self, message: str, level: str = "info") -> None:
        if self.dash is not None:
            self.dash.log(message, level)

    def _repair(self, message: str) -> None:
        """Append to the on-disk repair log so skips are auditable."""
        if self._repair_handle is None:
            return
        try:
            self._repair_handle.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')}\t{message}\n")
        except Exception:
            pass

    def _fail(self, context: str, exc: BaseException, level: str = "warn") -> None:
        self.stats.errors += 1
        text = f"{context}: {type(exc).__name__}: {exc}"
        self._log(text, level)
        self._repair(text)

    # ------------------------------------------------------------------
    # Main entrypoint
    # ------------------------------------------------------------------
    def run(self) -> Stats:
        stats = self.stats
        stats.started_at = time.monotonic()
        stats.pst_name = self.config.pst_path.name
        stats.backend = getattr(self.backend, "name", "unknown")

        if self._repair_log:
            try:
                self._repair_log.parent.mkdir(parents=True, exist_ok=True)
                self._repair_handle = self._repair_log.open("a", encoding="utf-8")
                self._repair(f"=== run started: {self.config.pst_path} ===")
            except OSError:
                self._repair_handle = None

        try:
            stats.phase = "opening"
            self._log(f"Opening {stats.pst_name} via {stats.backend}", "info")
            self.backend.open(self.config.pst_path)

            stats.phase = "indexing"
            folders = self._index_folders()

            if self.cancelled:
                self._log("Cancelled during indexing", "warn")
            else:
                stats.phase = "extracting"
                self._extract(folders)

        except KeyboardInterrupt:
            self._stop.set()
            stats.phase = "aborted"
            self._log("Interrupted by operator", "warn")
        except Exception as exc:  # backend-level fatal
            stats.phase = "failed"
            self._fail("FATAL during extraction", exc, level="error")
        finally:
            stats.phase = "done" if stats.phase != "failed" else "failed"
            stats.finished_at = time.monotonic()
            try:
                self.backend.close()
            except Exception as exc:
                self._repair(f"backend.close failed: {exc}")
            if self._repair_handle is not None:
                self._repair(
                    f"=== run finished: exported={stats.exported} "
                    f"dupes={stats.duplicates} spam={stats.spam} errors={stats.errors} ==="
                )
                try:
                    self._repair_handle.close()
                except Exception:
                    pass

        return stats

    # ------------------------------------------------------------------
    # Phase 1 — index
    # ------------------------------------------------------------------
    def _is_excluded(self, folder_name: str) -> bool:
        return folder_name.strip().lower() in self._excluded

    def _index_folders(self) -> list[tuple[object, str]]:
        """DFS the folder tree, building (folder_handle, full_path) pairs."""
        root = self.backend.root_folder()
        root_name = clean_field(self.backend.folder_name(root)) or "Root"

        discovered: list[tuple[object, str]] = []
        stack: list[tuple[object, str, int]] = [(root, root_name, 0)]
        guard = 0

        while stack and not self.cancelled:
            guard += 1
            if guard % 64 == 0:
                time.sleep(0)  # yield for the render thread
            folder, path, depth = stack.pop()

            self.stats.folders_discovered += 1
            self.stats.current_folder = path
            discovered.append((folder, path))

            if not self.config.include_subfolders or depth > 64:
                continue

            children: list[tuple[object, str, int]] = []
            try:
                for child in self.backend.subfolders(folder):
                    try:
                        name = clean_field(self.backend.folder_name(child)) or "(unnamed)"
                    except Exception as exc:
                        self._fail(f"folder name under {path}", exc)
                        continue
                    if self._is_excluded(name):
                        self._log(f"Skipping excluded folder: {path}\\{name}", "debug")
                        continue
                    children.append((child, f"{path}/{name}", depth + 1))
            except Exception as exc:
                self._fail(f"enumerating subfolders of {path}", exc)

            # Reverse-extend so DFS visits children in natural order.
            stack.extend(reversed(children))

            if self.stats.folders_discovered % 100 == 0:
                self._log(f"Indexed {self.stats.folders_discovered:,} folders…", "debug")

        self.stats.folders_total = len(discovered)

        # Count messages so phase 2 gets a determinate progress bar.
        total = 0
        for folder, path in discovered:
            try:
                total += self.backend.message_count(folder)
            except Exception:
                continue
            if total % 5000 < 10:
                time.sleep(0)
        self.stats.messages_total = total

        self._log(
            f"Index complete: {len(discovered):,} folders, {total:,} messages", "ok"
        )
        return discovered

    # ------------------------------------------------------------------
    # Phase 2 — extract
    # ------------------------------------------------------------------
    def _extract(self, folders: list[tuple[object, str]]) -> None:
        limit = self.config.limit
        start = time.monotonic()

        for folder, path in folders:
            if self.cancelled:
                break
            if limit and self.stats.exported >= limit:
                self._log(f"Reached export limit ({limit:,})", "ok")
                break

            self.stats.current_folder = path
            try:
                for message in self.backend.iter_messages(folder):
                    if self.cancelled:
                        break
                    self._handle_message(message, path)
                    if limit and self.stats.exported >= limit:
                        break
            except Exception as exc:
                self.stats.unreadable_folders += 1
                self._fail(f"folder read failed: {path}", exc)

            self.stats.folders_scanned += 1

            # Periodic heartbeat so the log is not silent on huge folders.
            if self.stats.folders_scanned % 50 == 0:
                self._log(
                    f"{self.stats.folders_scanned:,}/{self.stats.folders_total:,} folders · "
                    f"{self.stats.exported:,} exported",
                    "debug",
                )
                time.sleep(0)

        self.stats.bytes_written = self.writer.bytes_written
        self._log(
            f"Extraction complete in {time.monotonic() - start:,.1f}s", "ok"
        )

    def _handle_message(self, message: object, folder_path: str) -> None:
        self.stats.messages_seen += 1

        try:
            record = self.backend.read_message(message, folder_path)
        except Exception as exc:
            self._fail(f"unreadable message in {folder_path}", exc)
            return

        if self.config.max_body_chars:
            record.body_plain = clean_body(record.body_plain, limit=self.config.max_body_chars)

        # Spam first: cheaper than a dedupe hash and short-circuits.
        if self.spam_filter is not None:
            try:
                if self.spam_filter.is_spam(record, folder_path):
                    self.stats.spam += 1
                    return
            except Exception as exc:
                self._fail("spam filter", exc)

        if self.deduplicator is not None:
            try:
                if self.deduplicator.is_duplicate(record):
                    self.stats.duplicates += 1
                    return
            except Exception as exc:
                self._fail("deduplicator", exc)

        try:
            self.writer.write(record)
            self.stats.exported += 1
            self.stats.bytes_written = self.writer.bytes_written
        except Exception as exc:
            self._fail("csv write", exc)
