"""Interactive configuration wizard built on `questionary`."""
from __future__ import annotations

import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

import questionary
from questionary import ValidationError, Validator
from rich.console import Console
from rich.panel import Panel

from .errors import UserAbort
from .models import RunConfig
from .utils import human_size

console = Console()

STYLE = questionary.Style([
    ("qmark", "fg:#00afff bold"),
    ("question", "bold white"),
    ("answer", "fg:#00d787 bold"),
    ("pointer", "fg:#00afff bold"),
    ("highlighted", "fg:#00afff bold"),
    ("selected", "fg:#00d787"),
    ("separator", "fg:#6c6c6c"),
    ("instruction", "fg:#6c6c6c"),
    ("validation-error", "fg:#ff5f5f"),
])

_PST_SEARCH_ROOTS = (
    Path(os.environ.get("USERPROFILE", "")) / "Documents" / "Outlook Files",
    Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Outlook",
    Path(os.environ.get("USERPROFILE", "")) / "Documents",
    Path(os.environ.get("USERPROFILE", "")) / "Desktop",
    Path("C:/Backups"),
    Path("C:/PST"),
)

_OTHER = "▸ Other (type a path manually)"


# ----------------------------------------------------------------------
# Validators
# ----------------------------------------------------------------------
def _normalise(raw: str) -> Path:
    text = (raw or "").strip().strip('"').strip("'")
    return Path(os.path.expandvars(os.path.expanduser(text)))


class PSTFileValidator(Validator):
    def validate(self, document) -> None:
        path = _normalise(document.text)
        if not document.text.strip():
            raise ValidationError(message="A path is required", cursor_position=len(document.text))
        if not path.exists():
            raise ValidationError(message=f"Not found: {path}", cursor_position=len(document.text))
        if path.is_dir():
            raise ValidationError(message="That is a directory, not a file", cursor_position=len(document.text))
        if path.suffix.lower() not in (".pst", ".ost"):
            raise ValidationError(
                message=f"Expected .pst (got {path.suffix or 'no extension'})",
                cursor_position=len(document.text),
            )


class OutputPathValidator(Validator):
    def validate(self, document) -> None:
        if not document.text.strip():
            raise ValidationError(message="A path is required", cursor_position=len(document.text))
        path = _normalise(document.text)
        if path.suffix.lower() != ".csv":
            raise ValidationError(message="Output should end in .csv", cursor_position=len(document.text))
        parent = path.parent
        if not parent.exists():
            raise ValidationError(message=f"Directory does not exist: {parent}", cursor_position=len(document.text))
        if path.exists():
            try:
                if not os.access(path, os.W_OK):
                    raise ValidationError(message="File exists and is read-only", cursor_position=len(document.text))
            except OSError:
                pass


# ----------------------------------------------------------------------
# Discovery
# ----------------------------------------------------------------------
def discover_pst_files(limit: int = 20) -> list[Path]:
    found: list[Path] = []
    seen: set[str] = set()
    for root in _PST_SEARCH_ROOTS:
        try:
            if not root.is_dir():
                continue
            for pattern in ("*.pst", "*.ost"):
                for candidate in root.glob(pattern):
                    key = str(candidate).lower()
                    if key not in seen and candidate.is_file():
                        seen.add(key)
                        found.append(candidate)
        except OSError:
            continue
    found.sort(key=lambda p: p.stat().st_mtime if p.exists() else 0, reverse=True)
    return found[:limit]


# ----------------------------------------------------------------------
# Wizard
# ----------------------------------------------------------------------
def _ask_pst_path(preselect: Optional[Path] = None) -> Path:
    if preselect:
        return _normalise(str(preselect))

    candidates = discover_pst_files()
    if candidates:
        choices = [f"{p}   ({human_size(p.stat().st_size)})" for p in candidates]
        choices.append(_OTHER)
        answer = questionary.select(
            "Select the .pst file to process:",
            choices=choices,
            style=STYLE,
            use_shortcuts=False,
        ).ask()
        if answer is None:
            raise UserAbort()
        if answer != _OTHER:
            return Path(answer.split("   (")[0])
    else:
        console.print("[yellow]No .pst files auto-discovered in the usual locations.[/yellow]")

    answer = questionary.text(
        "Full path to the .pst file:",
        validate=PSTFileValidator,
        style=STYLE,
        instruction="(drag-and-drop works; quotes are stripped)",
    ).ask()
    if answer is None:
        raise UserAbort()
    return _normalise(answer)


def _default_output(pst_path: Path) -> str:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return str(pst_path.with_name(f"{pst_path.stem}_export_{stamp}.csv"))


def run_wizard(preselect_pst: Optional[Path] = None) -> RunConfig:
    """Collect a full RunConfig interactively."""
    console.print()
    console.print(
        Panel.fit(
            "[bold cyan]PST Export Toolkit[/bold cyan]\n"
            "[grey70]Extract Outlook mailboxes to a clean, streamed CSV[/grey70]",
            border_style="cyan",
            padding=(0, 3),
        )
    )
    console.print()

    # 1. Source ---------------------------------------------------------
    pst_path = _ask_pst_path(preselect_pst)
    size = pst_path.stat().st_size
    console.print(f"[green]✓[/green] Source: [white]{pst_path}[/white]  [grey50]({human_size(size)})[/grey50]\n")

    # 2. Output ---------------------------------------------------------
    output_answer = questionary.text(
        "Output CSV path:",
        default=_default_output(pst_path),
        validate=OutputPathValidator,
        style=STYLE,
    ).ask()
    if output_answer is None:
        raise UserAbort()
    output_path = _normalise(output_answer)
    console.print(f"[green]✓[/green] Output: [white]{output_path}[/white]\n")

    # 3. Duplicates -----------------------------------------------------
    remove_duplicates = questionary.confirm(
        "Remove duplicate messages?",
        default=True,
        style=STYLE,
    ).ask()
    if remove_duplicates is None:
        raise UserAbort()

    dedupe_mode = "message_id"
    if remove_duplicates:
        dedupe_mode = questionary.select(
            "Deduplication strategy:",
            choices=[
                questionary.Choice(
                    "Message-ID header  (exact, cheapest)",
                    value="message_id",
                ),
                questionary.Choice(
                    "Content fingerprint  (catches re-sent copies)",
                    value="content",
                ),
            ],
            default="message_id",
            style=STYLE,
        ).ask()
        if dedupe_mode is None:
            raise UserAbort()

    # 4. Spam -----------------------------------------------------------
    filter_spam = questionary.confirm(
        "Filter out spam / junk mail?",
        default=True,
        style=STYLE,
    ).ask()
    if filter_spam is None:
        raise UserAbort()
    console.print(
        "[grey50]  Excludes Junk E-mail folders and messages carrying spam headers.[/grey50]\n"
    )

    # 5. Advanced -------------------------------------------------------
    advanced = questionary.confirm("Configure advanced options?", default=False, style=STYLE).ask()
    if advanced is None:
        raise UserAbort()

    include_subfolders = True
    exclude_folders: list[str] = []
    max_body_chars = 0
    csv_delimiter = ","
    backend = "auto"

    if advanced:
        include_subfolders = bool(
            questionary.confirm("Recurse into subfolders?", default=True, style=STYLE).ask()
        )

        exclude_answer = questionary.text(
            "Folder names to skip (comma-separated, blank for none):",
            default="Deleted Items, Sync Issues, RSS Feeds",
            style=STYLE,
        ).ask()
        if exclude_answer:
            exclude_folders = [
                part.strip() for part in exclude_answer.split(",") if part.strip()
            ]

        body_answer = questionary.text(
            "Truncate each body to N characters (0 = unlimited):",
            default="0",
            style=STYLE,
            validate=lambda v: v.strip().isdigit() or "Enter a whole number",
        ).ask()
        max_body_chars = int((body_answer or "0").strip())

        delimiter = questionary.select(
            "CSV delimiter:",
            choices=[
                questionary.Choice("Comma  ,", value=","),
                questionary.Choice("Semicolon  ;  (European Excel)", value=";"),
                questionary.Choice("Tab", value="\t"),
            ],
            style=STYLE,
        ).ask()
        csv_delimiter = delimiter or ","

        backend = questionary.select(
            "PST reader backend:",
            choices=[
                questionary.Choice("Auto-detect (recommended)", value="auto"),
                questionary.Choice("libpff / pypff", value="pypff"),
                questionary.Choice("Outlook COM", value="outlook"),
            ],
            style=STYLE,
        ).ask() or "auto"

    config = RunConfig(
        pst_path=pst_path,
        output_path=output_path,
        remove_duplicates=bool(remove_duplicates),
        dedupe_mode=dedupe_mode,
        filter_spam=bool(filter_spam),
        include_subfolders=include_subfolders,
        exclude_folders=exclude_folders,
        backend=backend,
        max_body_chars=max_body_chars,
        csv_delimiter=csv_delimiter,
        log_path=output_path.with_suffix(".repair.log"),
    )

    _print_summary(config)
    confirmed = questionary.confirm("Start extraction?", default=True, style=STYLE).ask()
    if not confirmed:
        raise UserAbort()

    return config


def _print_summary(config: RunConfig) -> None:
    from rich.table import Table

    table = Table(show_header=False, box=None, padding=(0, 2))
    table.add_column(style="grey70", justify="right")
    table.add_column(style="white")

    table.add_row("Source", str(config.pst_path))
    table.add_row("Output", str(config.output_path))
    table.add_row("Deduplicate", f"{'yes' if config.remove_duplicates else 'no'}"
                                 + (f" ({config.dedupe_mode})" if config.remove_duplicates else ""))
    table.add_row("Filter spam", "yes" if config.filter_spam else "no")
    table.add_row("Recurse", "yes" if config.include_subfolders else "no")
    table.add_row("Excluded", ", ".join(config.exclude_folders) or "—")
    table.add_row("Backend", config.backend)
    table.add_row("Body cap", f"{config.max_body_chars:,}" if config.max_body_chars else "unlimited")

    console.print()
    console.print(Panel(table, title="[bold]Run Summary[/bold]", border_style="blue", box=None))
    console.print()


def is_interactive() -> bool:
    try:
        return sys.stdin is not None and sys.stdin.isatty()
    except Exception:
        return False
