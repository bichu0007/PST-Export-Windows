#!/usr/bin/env python
"""PST Export Toolkit — Windows CLI entrypoint.

Extracts Outlook .pst mail into a streamed, deduplicated CSV with a live
rich dashboard.
"""
from __future__ import annotations

import argparse
import sys
import threading
from datetime import datetime
from pathlib import Path

if sys.platform == "win32":
    try:
        import colorama  # noqa: F401  (rich handles most of this, but ANSI on cmd.exe)
    except Exception:
        pass

from rich import box
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

from pstexport import __version__
from pstexport.backends import available_backends, load_backend
from pstexport.config import is_interactive, run_wizard
from pstexport.errors import BackendUnavailableError, OutputError, UserAbort
from pstexport.extractor import ExtractionEngine
from pstexport.models import RunConfig, Stats
from pstexport.repair import find_scanpst, launch_scanpst, probe_pst
from pstexport.ui import Dashboard
from pstexport.utils import human_duration, human_size
from pstexport.writers import CsvStreamWriter, NullWriter

console = Console()


# ----------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pst-export",
        description="Extract Outlook .pst mailboxes to CSV with live progress.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  pst-export\n"
            "  pst-export --pst mail.pst --out mail.csv --yes\n"
            "  pst-export --pst mail.pst --no-dedupe --no-spam-filter --backend outlook\n"
        ),
    )
    parser.add_argument("--pst", type=Path, help="Path to the .pst file.")
    parser.add_argument("--out", type=Path, help="Destination .csv path.")
    parser.add_argument("--yes", "-y", action="store_true", help="Skip the wizard and use flags/defaults.")
    parser.add_argument("--no-dedupe", action="store_true", help="Disable duplicate removal.")
    parser.add_argument("--dedupe-mode", choices=("message_id", "content"), default="message_id")
    parser.add_argument("--no-spam-filter", action="store_true", help="Keep junk/spam messages.")
    parser.add_argument("--exclude", action="append", default=[], metavar="FOLDER",
                        help="Folder name to skip. Repeatable.")
    parser.add_argument("--no-recurse", action="store_true", help="Top-level folders only.")
    parser.add_argument("--backend", choices=("auto", "pypff", "outlook"), default="auto")
    parser.add_argument("--max-body", type=int, default=0, help="Truncate bodies to N chars (0=off).")
    parser.add_argument("--limit", type=int, default=0, help="Stop after N exported rows.")
    parser.add_argument("--delimiter", default=",", help="CSV delimiter (default ',').")
    parser.add_argument("--no-tui", action="store_true", help="Plain scrolling output instead of the dashboard.")
    parser.add_argument("--dry-run", action="store_true", help="Parse everything, write nothing.")
    parser.add_argument("--repair", action="store_true", help="Launch SCANPST.EXE before extracting.")
    parser.add_argument("--version", action="version", version=f"pst-export {__version__}")
    return parser


def config_from_args(args: argparse.Namespace) -> RunConfig:
    pst_path: Path = args.pst
    output = args.out or pst_path.with_name(
        f"{pst_path.stem}_export_{datetime.now():%Y%m%d_%H%M%S}.csv"
    )
    return RunConfig(
        pst_path=pst_path,
        output_path=output,
        remove_duplicates=not args.no_dedupe,
        dedupe_mode=args.dedupe_mode,
        filter_spam=not args.no_spam_filter,
        include_subfolders=not args.no_recurse,
        exclude_folders=list(args.exclude),
        backend=args.backend,
        max_body_chars=max(0, args.max_body),
        limit=max(0, args.limit),
        csv_delimiter=args.delimiter,
        dry_run=args.dry_run,
        log_path=output.with_suffix(".repair.log"),
    )


# ----------------------------------------------------------------------
# Preflight
# ----------------------------------------------------------------------
def preflight(config: RunConfig, interactive: bool, force_repair: bool) -> bool:
    probe = probe_pst(config.pst_path, config.output_path.parent)

    if not probe.usable:
        console.print(Panel(f"[bold red]{probe.fatal}[/bold red]", title="Preflight failed", border_style="red"))
        return False

    if probe.warnings:
        body = "\n".join(f"• {w}" for w in probe.warnings)
        console.print(Panel(body, title="[bold yellow]Warnings[/bold yellow]", border_style="yellow"))

    if probe.appears_locked:
        console.print(
            "[yellow]The file appears locked. Reading usually still works; "
            "close Outlook if the run fails.[/yellow]"
        )

    should_repair = force_repair or (probe.warnings and not probe.magic_ok)

    if should_repair:
        exe = find_scanpst()
        if exe is None:
            console.print("[yellow]SCANPST.EXE not found — continuing without repair.[/yellow]")
        else:
            console.print(f"[cyan]Found repair tool:[/cyan] {exe}")
            if not interactive or args_confirm("Run SCANPST.EXE now? (Outlook must be closed)"):
                launched, message = launch_scanpst(config.pst_path, wait=False)
                console.print(f"[cyan]{message}[/cyan]")
                if launched and interactive:
                    args_confirm("Press Enter once SCANPST.EXE has finished…", default=True)

    return True


def args_confirm(prompt: str, default: bool = True) -> bool:
    suffix = "[Y/n]" if default else "[y/N]"
    try:
        answer = input(f"{prompt} {suffix} ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        return default
    if not answer:
        return default
    return answer in ("y", "yes")


# ----------------------------------------------------------------------
# Reporting
# ----------------------------------------------------------------------
def print_summary(stats: Stats, config: RunConfig, output_path: Path) -> None:
    table = Table(show_header=False, box=box.SIMPLE, padding=(0, 2))
    table.add_column(style="grey70", justify="right")
    table.add_column(style="bold white")

    table.add_row("Source", str(config.pst_path))
    table.add_row("Backend", stats.backend)
    table.add_row("Folders scanned", f"{stats.folders_scanned:,} / {stats.folders_total:,}")
    table.add_row("Messages seen", f"{stats.messages_seen:,}")
    table.add_row("Exported", f"[green]{stats.exported:,}[/green]")
    table.add_row("Duplicates skipped", f"[yellow]{stats.duplicates:,}[/yellow]")
    table.add_row("Spam filtered", f"[magenta]{stats.spam:,}[/magenta]")
    table.add_row("Errors / skips", f"[red]{stats.errors:,}[/red]" if stats.errors else "0")
    table.add_row("Unreadable folders", f"{stats.unreadable_folders:,}")
    table.add_row("Output size", human_size(stats.bytes_written))
    table.add_row("Duration", human_duration(stats.elapsed))
    table.add_row("Throughput", f"{stats.rate:,.0f} msg/s")

    console.print()
    console.print(Panel(table, title="[bold]Extraction Report[/bold]", border_style="green"))

    if stats.errors:
        console.print(f"[yellow]Repair log:[/yellow] {config.log_path}")
    if not config.dry_run and stats.exported:
        console.print(f"[green]CSV written:[/green] {output_path}")
    console.print()


# ----------------------------------------------------------------------
# Entrypoint
# ----------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    global args
    parser = build_parser()
    args = parser.parse_args(argv)

    interactive = is_interactive() and not args.no_tui

    # --- backend availability gate ------------------------------------
    usable = available_backends()
    if not usable:
        console.print(Panel(
            "[bold red]No PST reader available.[/bold red]\n\n"
            "Install one of:\n"
            "  [cyan]pip install libpff-python[/cyan]   (recommended, no Outlook needed)\n"
            "  [cyan]pip install pywin32[/cyan]          (requires Outlook installed)",
            title="Missing dependency",
            border_style="red",
        ))
        return 2

    # --- gather configuration -----------------------------------------
    try:
        if args.yes and args.pst:
            config = config_from_args(args)
        else:
            if not interactive:
                if not args.pst:
                    console.print("[red]--pst is required in non-interactive mode.[/red]")
                    return 2
                config = config_from_args(args)
            else:
                config = run_wizard(args.pst)
    except UserAbort:
        console.print("\n[yellow]Aborted by user.[/yellow]")
        return 130
    except KeyboardInterrupt:
        console.print("\n[yellow]Interrupted.[/yellow]")
        return 130

    # --- preflight -----------------------------------------------------
    if not preflight(config, interactive, args.repair):
        return 1

    # --- backend -------------------------------------------------------
    try:
        backend = load_backend(config.backend)
    except BackendUnavailableError as exc:
        console.print(f"[red]{exc}[/red]")
        return 2

    console.print(f"[cyan]Using backend:[/cyan] {backend.name} — {backend.description}")

    # --- writer --------------------------------------------------------
    stats = Stats()
    writer = NullWriter()
    try:
        if not config.dry_run:
            writer = CsvStreamWriter(
                config.output_path,
                encoding=config.csv_encoding,
                delimiter=config.csv_delimiter,
                excel_safe_cells=True,
            )
    except OutputError as exc:
        console.print(f"[red]{exc}[/red]")
        return 1

    engine = ExtractionEngine(config, backend, writer, stats)

    # --- run with dashboard -------------------------------------------
    dashboard = Dashboard(
        stats,
        title="PST Export",
        enabled=not args.no_tui,
        console=console,
    )
    engine.dash = dashboard

    exit_code = 0
    try:
        with dashboard:
            worker = threading.Thread(target=engine.run, name="extractor", daemon=True)
            worker.start()
            try:
                while worker.is_alive():
                    worker.join(0.25)
            except KeyboardInterrupt:
                dashboard.log("Ctrl+C — finishing current folder…", "warn")
                engine.request_stop()
                worker.join(30)
        if stats.phase == "failed":
            exit_code = 1
    finally:
        try:
            writer.close()
        except OutputError as exc:
            console.print(f"[red]{exc}[/red]")
            exit_code = 1

    print_summary(stats, config, config.output_path)

    if interactive and not config.dry_run and stats.exported and args_confirm("Open output folder?"):
        import subprocess
        try:
            subprocess.Popen(["explorer", "/select,", str(config.output_path)])
        except OSError:
            pass

    return exit_code


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit(130)
