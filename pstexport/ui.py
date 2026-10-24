"""Live `rich` dashboard. Renders on the main thread at a fixed refresh rate."""
from __future__ import annotations

import threading
import time
from collections import deque
from typing import Deque, Optional

from rich import box
from rich.console import Console, Group
from rich.layout import Layout
from rich.live import Live
from rich.markup import escape
from rich.panel import Panel
from rich.progress_bar import ProgressBar
from rich.table import Table
from rich.text import Text

from .models import Stats
from .utils import human_duration, human_size, truncate_middle

_LEVEL_STYLE = {
    "debug": "grey50",
    "info": "cyan",
    "ok": "green",
    "warn": "yellow",
    "error": "bold red",
}

_VISIBLE_LOG_LINES = 14


class Dashboard:
    """Full-screen Live dashboard.

    The extraction engine calls `log()` from its worker thread; `rich.Live`
    invokes `_render()` from the main thread. All shared state is lock-guarded.
    """

    def __init__(
        self,
        stats: Stats,
        title: str = "PST Export",
        max_log_lines: int = 400,
        refresh_per_second: int = 12,
        enabled: bool = True,
        console: Optional[Console] = None,
    ) -> None:
        self.stats = stats
        self.title = title
        self.enabled = enabled
        self.console = console or Console()

        self._lock = threading.RLock()
        self._log: Deque[tuple[float, str, str]] = deque(maxlen=max_log_lines)
        self._live: Optional[Live] = None
        self._refresh = max(1, refresh_per_second)
        self._frames = 0

    # ------------------------------------------------------------------
    # Public API (worker-thread safe)
    # ------------------------------------------------------------------
    def log(self, message: str, level: str = "info") -> None:
        with self._lock:
            self._log.append((time.time(), level, message))
        if not self.enabled:
            stamp = time.strftime("%H:%M:%S")
            prefix = {"warn": "WARN ", "error": "ERROR", "ok": "OK   "}.get(level, "     ")
            self.console.print(f"[grey50]{stamp}[/grey50] {prefix} {escape(message)}")

    def start(self) -> None:
        if not self.enabled:
            return
        self._live = Live(
            self._render,
            console=self.console,
            refresh_per_second=self._refresh,
            screen=True,
            transient=False,
            auto_refresh=True,
        )
        self._live.start()

    def stop(self) -> None:
        if self._live is not None:
            try:
                self._live.stop()
            except Exception:
                pass
            self._live = None

    def __enter__(self) -> "Dashboard":
        self.start()
        return self

    def __exit__(self, *exc) -> None:
        self.stop()

    # ------------------------------------------------------------------
    # Rendering (main-thread only)
    # ------------------------------------------------------------------
    def _render(self) -> Layout:
        self._frames += 1
        layout = Layout()
        layout.split_column(
            Layout(self._header(), name="header", size=3),
            Layout(self._progress_panel(), name="progress", size=5),
            Layout(self._counters(), name="counters", size=5),
            Layout(self._current_folder(), name="current", size=3),
            Layout(self._log_panel(), name="log"),
        )
        return layout

    def _header(self) -> Panel:
        stats = self.stats
        spinner = "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏"[self._frames % 10]

        left = Text()
        left.append(f"{spinner} ", style="bold cyan")
        left.append(f"{self.title}", style="bold white")
        left.append(f"  ·  {stats.phase.upper()}", style="bold yellow")
        left.append(f"  ·  {stats.backend or 'n/a'}", style="grey50")

        right = Text(justify="right")
        right.append("elapsed ", style="grey50")
        right.append(human_duration(stats.elapsed), style="white")
        right.append("   eta ", style="grey50")
        right.append(human_duration(stats.eta), style="white")

        table = Table.grid(expand=True)
        table.add_column(ratio=3)
        table.add_column(ratio=2, justify="right")
        table.add_row(left, right)
        return Panel(table, border_style="cyan", box=box.ROUNDED, padding=(0, 1))

    def _progress_panel(self) -> Panel:
        stats = self.stats
        total = stats.messages_total or stats.folders_total or 1
        done = stats.messages_seen if stats.messages_total else stats.folders_scanned
        percent = min(100.0, (done / total) * 100) if total else 0.0

        bar = ProgressBar(
            total=max(1, total),
            completed=min(done, total),
            width=None,
            complete_style="green",
            finished_style="bright_green",
            pulse_style="cyan",
            bar_style="grey30",
        )

        caption = Text()
        caption.append(f"{percent:5.1f}%  ", style="bold green")
        if stats.messages_total:
            caption.append(
                f"{done:,} / {stats.messages_total:,} messages", style="white"
            )
        else:
            caption.append(
                f"{stats.folders_scanned:,} / {stats.folders_total:,} folders", style="white"
            )
        caption.append("   ·   ", style="grey50")
        caption.append(f"{stats.rate:,.0f} msg/s", style="cyan")

        return Panel(
            Group(bar, caption),
            title="[bold]Progress[/bold]",
            border_style="green",
            box=box.ROUNDED,
            padding=(0, 1),
        )

    def _counters(self) -> Panel:
        stats = self.stats
        grid = Table.grid(expand=True)
        for _ in range(5):
            grid.add_column(justify="center", ratio=1)

        grid.add_row(
            Text("PROCESSED", style="bold grey70"),
            Text("EXPORTED", style="bold grey70"),
            Text("DUPLICATES", style="bold grey70"),
            Text("SPAM", style="bold grey70"),
            Text("ERRORS", style="bold grey70"),
        )
        grid.add_row(
            Text(f"{stats.messages_seen:,}", style="bold white"),
            Text(f"{stats.exported:,}", style="bold bright_green"),
            Text(f"{stats.duplicates:,}", style="bold yellow"),
            Text(f"{stats.spam:,}", style="bold magenta"),
            Text(f"{stats.errors:,}", style="bold red" if stats.errors else "grey50"),
        )
        grid.add_row(
            Text(f"folders {stats.folders_scanned:,}", style="grey50"),
            Text(human_size(stats.bytes_written), style="grey50"),
            Text(f"{stats.duplicates / max(1, stats.messages_seen) * 100:.1f}%", style="grey50"),
            Text(f"unreadable {stats.unreadable_folders:,}", style="grey50"),
            Text("fatal: no" if stats.errors < 500 else "fatal: check log", style="grey50"),
        )
        return Panel(grid, title="[bold]Counters[/bold]", border_style="blue", box=box.ROUNDED, padding=(0, 1))

    def _current_folder(self) -> Panel:
        folder = self.stats.current_folder or "—"
        width = max(20, self.console.width - 24)
        text = Text(truncate_middle(folder, width), style="white", overflow="ellipsis")
        return Panel(
            text,
            title="[bold]Current Folder[/bold]",
            border_style="magenta",
            box=box.ROUNDED,
            padding=(0, 1),
        )

    def _log_panel(self) -> Panel:
        with self._lock:
            entries = list(self._log)[-_VISIBLE_LOG_LINES:]

        if not entries:
            body: Text | Group = Text("waiting for first folder…", style="grey50")
        else:
            lines = Text()
            for timestamp, level, message in entries:
                stamp = time.strftime("%H:%M:%S", time.localtime(timestamp))
                style = _LEVEL_STYLE.get(level, "white")
                lines.append(f"{stamp} ", style="grey50")
                lines.append(f"{level.upper():<5} ", style=f"bold {style}")
                lines.append(f"{message}\n", style=style)
            body = lines

        return Panel(
            body,
            title="[bold]Activity Log[/bold]",
            border_style="grey37",
            box=box.ROUNDED,
            padding=(0, 1),
        )
