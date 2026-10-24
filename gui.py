#!/usr/bin/env python
"""Native Tkinter GUI for PST Export Toolkit.

The GUI calls the existing PST backend/extraction engine. All extraction work
runs on a worker thread; Tk widgets are only updated by the main thread.
"""
from __future__ import annotations

import logging
import queue
import re
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Any

from pstexport.backends import available_backends, load_backend
from pstexport.errors import BackendUnavailableError, OutputError
from pstexport.extractor import ExtractionEngine
from pstexport.models import RunConfig, Stats
from pstexport.repair import find_scanpst, launch_scanpst, probe_pst
from pstexport.utils import human_duration, human_size
from pstexport.writers import CsvStreamWriter

APP_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = APP_DIR / "extracted_output"
LOG_DIR = APP_DIR / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    filename=LOG_DIR / "gui.log",
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(threadName)s %(message)s",
    encoding="utf-8",
)


class QueueLogSink:
    """Small adapter implementing the ``dashboard.log`` interface."""

    def __init__(self, events: queue.Queue) -> None:
        self.events = events

    def log(self, message: str, level: str = "info") -> None:
        self.events.put(("log", message, level))


def readable_name(path: Path, source_root: Path | None) -> str:
    """Make a safe, reasonably unique output stem for single or batch mode."""
    try:
        relative = path.relative_to(source_root) if source_root and source_root.is_dir() else Path(path.name)
    except ValueError:
        relative = Path(path.name)
    stem = "__".join(relative.with_suffix("").parts)
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._-")
    return stem or "pst_export"


def unique_output_path(folder: Path, stem: str) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = folder / f"{stem}_export_{stamp}.csv"
    candidate = base
    suffix = 2
    while candidate.exists():
        candidate = folder / f"{stem}_export_{stamp}_{suffix}.csv"
        suffix += 1
    return candidate


class PSTExportGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PST Export Toolkit")
        self.geometry("1050x760")
        self.minsize(850, 630)

        try:
            self.tk.call("tk", "scaling", 1.1)
            style = ttk.Style(self)
            if "vista" in style.theme_names():
                style.theme_use("vista")
            elif "clam" in style.theme_names():
                style.theme_use("clam")
            style.configure("Title.TLabel", font=("Segoe UI", 19, "bold"))
            style.configure("Subtitle.TLabel", font=("Segoe UI", 10))
            style.configure("Start.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8))
            style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
            style.configure("Metric.TLabel", font=("Segoe UI", 10, "bold"))
        except tk.TclError:
            pass

        self.events: queue.Queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.active_engine: ExtractionEngine | None = None
        self.active_stats: Stats | None = None
        self.active_index = 0
        self.total_files = 0
        self.completed_files = 0
        self.running = False
        self.pending_repair_ack: threading.Event | None = None
        self._last_progress_phase = ""

        try:
            DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            initial_output = str(DEFAULT_OUTPUT_DIR)
        except OSError:
            fallback = Path.home() / "Documents" / "PST Export Output"
            fallback.mkdir(parents=True, exist_ok=True)
            initial_output = str(fallback)

        self.source_var = tk.StringVar()
        self.output_var = tk.StringVar(value=initial_output)
        self.source_kind = tk.StringVar(value="file")
        self.recursive_var = tk.BooleanVar(value=True)
        self.repair_var = tk.BooleanVar(value=False)
        self.dedupe_var = tk.BooleanVar(value=True)
        self.spam_var = tk.BooleanVar(value=True)
        self.backend_var = tk.StringVar(value="auto")
        self.status_var = tk.StringVar(value="Ready")
        self.current_var = tk.StringVar(value="Choose a PST file or a folder containing PST files.")
        self.progress_var = tk.DoubleVar(value=0.0)
        self.percent_var = tk.StringVar(value="0%")
        self.metrics_var = tk.StringVar(value="Messages: 0 / 0    Exported: 0    Duplicates: 0    Spam filtered: 0    Errors: 0")
        self._build_ui()
        self._append_log("Ready. Select a PST file or a folder to begin.", "info")
        self.after(120, self._poll_events)
        self.after(250, self._refresh_progress)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=(18, 15, 18, 12))
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 13))
        ttk.Label(header, text="PST Export Toolkit", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Extract Outlook mailboxes to Excel-friendly CSV files",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        source_box = ttk.LabelFrame(outer, text="  1. Source files  ", padding=12, style="Section.TLabelframe")
        source_box.pack(fill="x", pady=(0, 9))
        source_line = ttk.Frame(source_box)
        source_line.pack(fill="x")
        self.source_entry = ttk.Entry(source_line, textvariable=self.source_var)
        self.source_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.browse_file_button = ttk.Button(source_line, text="Browse PST File…", command=self._browse_file)
        self.browse_file_button.pack(side="left", padx=(0, 6))
        self.browse_folder_button = ttk.Button(source_line, text="Browse Folder…", command=self._browse_folder)
        self.browse_folder_button.pack(side="left")
        ttk.Label(
            source_box,
            text="Choose one .pst file or a folder. Folder mode processes matching PST files into separate CSVs.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(7, 0))

        output_box = ttk.LabelFrame(outer, text="  2. Output location  ", padding=12, style="Section.TLabelframe")
        output_box.pack(fill="x", pady=(0, 9))
        output_line = ttk.Frame(output_box)
        output_line.pack(fill="x")
        self.output_entry = ttk.Entry(output_line, textvariable=self.output_var)
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.browse_output_button = ttk.Button(output_line, text="Browse Output…", command=self._browse_output)
        self.browse_output_button.pack(side="left")
        ttk.Label(
            output_box,
            text="The default extracted_output folder is created automatically. Each PST produces its own CSV and repair log.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(7, 0))

        options_box = ttk.LabelFrame(outer, text="  3. Processing options  ", padding=10, style="Section.TLabelframe")
        options_box.pack(fill="x", pady=(0, 10))
        options_grid = ttk.Frame(options_box)
        options_grid.pack(fill="x")
        self.repair_check = ttk.Checkbutton(options_grid, text="Scan/repair with SCANPST before extraction", variable=self.repair_var)
        self.repair_check.grid(row=0, column=0, sticky="w", padx=(0, 20), pady=3)
        self.recursive_check = ttk.Checkbutton(options_grid, text="Search subfolders in folder mode", variable=self.recursive_var)
        self.recursive_check.grid(row=0, column=1, sticky="w", pady=3)
        self.dedupe_check = ttk.Checkbutton(options_grid, text="Remove duplicates", variable=self.dedupe_var)
        self.dedupe_check.grid(row=1, column=0, sticky="w", padx=(0, 20), pady=3)
        self.spam_check = ttk.Checkbutton(options_grid, text="Filter likely spam/junk", variable=self.spam_var)
        self.spam_check.grid(row=1, column=1, sticky="w", pady=3)
        ttk.Label(options_grid, text="PST reader:").grid(row=2, column=0, sticky="w", pady=(7, 0))
        self.backend_combo = ttk.Combobox(options_grid, textvariable=self.backend_var, values=("auto", "pypff", "outlook"), state="readonly", width=14)
        self.backend_combo.grid(row=2, column=0, sticky="w", padx=(85, 0), pady=(7, 0))
        ttk.Label(options_grid, text="Auto tries the native reader first, then Outlook.", style="Subtitle.TLabel").grid(row=2, column=1, sticky="w", pady=(7, 0))

        action_line = ttk.Frame(outer)
        action_line.pack(fill="x", pady=(0, 10))
        self.start_button = ttk.Button(action_line, text="Start Extraction", style="Start.TButton", command=self._start)
        self.start_button.pack(side="left")
        self.cancel_button = ttk.Button(action_line, text="Cancel", command=self._cancel, state="disabled")
        self.cancel_button.pack(side="left", padx=(8, 0))
        self.open_output_button = ttk.Button(action_line, text="Open Output Folder", command=self._open_output)
        self.open_output_button.pack(side="right")

        progress_box = ttk.LabelFrame(outer, text="  Live status  ", padding=10, style="Section.TLabelframe")
        progress_box.pack(fill="x", pady=(0, 10))
        status_line = ttk.Frame(progress_box)
        status_line.pack(fill="x")
        ttk.Label(status_line, textvariable=self.status_var, style="Metric.TLabel").pack(side="left")
        ttk.Label(status_line, textvariable=self.percent_var, style="Metric.TLabel").pack(side="right")
        self.progress_bar = ttk.Progressbar(progress_box, variable=self.progress_var, maximum=100, mode="determinate")
        self.progress_bar.pack(fill="x", pady=(7, 6))
        ttk.Label(progress_box, textvariable=self.current_var).pack(anchor="w")
        ttk.Label(progress_box, textvariable=self.metrics_var, style="Subtitle.TLabel").pack(anchor="w", pady=(6, 0))

        log_box = ttk.LabelFrame(outer, text="  Activity log  ", padding=8, style="Section.TLabelframe")
        log_box.pack(fill="both", expand=True)
        self.log_text = scrolledtext.ScrolledText(
            log_box, height=12, wrap="word", state="disabled", font=("Consolas", 9),
            background="#101820", foreground="#e7edf2", insertbackground="#ffffff",
            relief="flat", padx=10, pady=8,
        )
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_configure("info", foreground="#b9d7f5")
        self.log_text.tag_configure("ok", foreground="#8ce0ad")
        self.log_text.tag_configure("warn", foreground="#ffd166")
        self.log_text.tag_configure("error", foreground="#ff8585")
        self.log_text.tag_configure("debug", foreground="#99a5b1")

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(7, 0))
        ttk.Label(footer, text="Logs: logs/gui.log   •   CSV files: selected output folder", style="Subtitle.TLabel").pack(anchor="w")

    def _browse_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Select an Outlook PST file",
            filetypes=(("Outlook PST files", "*.pst"), ("All files", "*.*")),
        )
        if path:
            self.source_var.set(path)
            self.source_kind.set("file")

    def _browse_folder(self) -> None:
        path = filedialog.askdirectory(title="Select a folder containing PST files", mustexist=True)
        if path:
            self.source_var.set(path)
            self.source_kind.set("folder")

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(title="Choose output folder", mustexist=False)
        if path:
            self.output_var.set(path)

    def _discover_sources(self) -> tuple[list[Path], Path | None]:
        raw = self.source_var.get().strip().strip('"')
        if not raw:
            raise ValueError("Choose a PST file or a folder first.")
        source = Path(raw).expanduser().resolve()
        if source.is_file():
            if source.suffix.lower() != ".pst":
                raise ValueError("The selected file does not have a .pst extension.")
            return [source], source.parent
        if source.is_dir():
            iterator = source.rglob("*.pst") if self.recursive_var.get() else source.glob("*.pst")
            found = sorted((p for p in iterator if p.is_file()), key=lambda p: str(p).lower())
            if not found:
                raise ValueError("No .pst files were found in the selected folder.")
            return found, source
        raise ValueError(f"Source path does not exist: {source}")

    def _start(self) -> None:
        if self.running:
            return
        try:
            paths, source_root = self._discover_sources()
            output_dir = Path(self.output_var.get().strip().strip('"')).expanduser().resolve()
            if not self.output_var.get().strip():
                raise ValueError("Choose an output folder.")
            output_dir.mkdir(parents=True, exist_ok=True)
            if not output_dir.is_dir():
                raise ValueError("Output path must be a directory.")
        except (OSError, ValueError) as exc:
            messagebox.showerror("Check your paths", str(exc), parent=self)
            return

        readers = available_backends()
        selected_backend = self.backend_var.get()
        if selected_backend == "auto" and not readers:
            messagebox.showerror(
                "PST reader not installed",
                "No supported PST reader is available. Install a compatible libpff-python package, or install Microsoft Outlook and pywin32.\n\nUse the launcher to install the standard app dependencies first.",
                parent=self,
            )
            return
        if selected_backend != "auto" and selected_backend not in readers:
            messagebox.showerror(
                "Selected reader unavailable",
                f"The '{selected_backend}' reader is not available in this Python environment.\n\nAvailable readers: {', '.join(readers) or 'none'}",
                parent=self,
            )
            return

        self.cancel_event.clear()
        self.active_engine = None
        self.active_stats = None
        self.active_index = 0
        self.total_files = len(paths)
        self.completed_files = 0
        self.running = True
        self.progress_var.set(0)
        self.percent_var.set("0%")
        self.status_var.set("Preparing…")
        self.current_var.set(f"Queued {len(paths)} PST file(s).")
        self._set_controls_running(True)
        self._append_log(f"Queued {len(paths)} PST file(s). Output: {output_dir}", "info")

        options = {
            "output_dir": output_dir,
            "source_root": source_root,
            "backend": selected_backend,
            "repair": self.repair_var.get(),
            "dedupe": self.dedupe_var.get(),
            "spam": self.spam_var.get(),
        }
        self.worker = threading.Thread(target=self._run_job, args=(paths, options), name="PST-GUI-Worker", daemon=True)
        self.worker.start()

    def _run_job(self, paths: list[Path], options: dict[str, Any]) -> None:
        results: list[dict[str, Any]] = []
        output_dir: Path = options["output_dir"]
        source_root: Path | None = options["source_root"]
        try:
            for index, pst_path in enumerate(paths, start=1):
                if self.cancel_event.is_set():
                    break
                self.events.put(("file_start", index, len(paths), pst_path))
                self.events.put(("log", f"[{index}/{len(paths)}] Preparing {pst_path}", "info"))

                if options["repair"]:
                    acknowledged = threading.Event()
                    outcome: dict[str, Any] = {"continue": True}
                    self.events.put(("repair_request", pst_path, acknowledged, outcome))
                    while not acknowledged.wait(0.2):
                        if self.cancel_event.is_set():
                            outcome["continue"] = False
                            acknowledged.set()
                    if not outcome.get("continue", False) or self.cancel_event.is_set():
                        self.events.put(("log", f"Repair stage cancelled for {pst_path.name}.", "warn"))
                        break

                probe = probe_pst(pst_path, output_dir)
                if not probe.usable:
                    self.events.put(("log", f"Skipping {pst_path.name}: {probe.fatal}", "error"))
                    results.append({"path": pst_path, "output": None, "stats": None, "failed": True})
                    self.events.put(("file_complete", pst_path, None, None, True))
                    continue
                for warning in probe.warnings:
                    self.events.put(("log", f"{pst_path.name}: {warning}", "warn"))

                stem = readable_name(pst_path, source_root if source_root and source_root.is_dir() else None)
                output_path = unique_output_path(output_dir, stem)
                config = RunConfig(
                    pst_path=pst_path,
                    output_path=output_path,
                    remove_duplicates=bool(options["dedupe"]),
                    filter_spam=bool(options["spam"]),
                    include_subfolders=True,
                    backend=str(options["backend"]),
                    log_path=output_path.with_suffix(".repair.log"),
                )
                stats = Stats()
                writer = None
                try:
                    backend = load_backend(config.backend)
                    writer = CsvStreamWriter(
                        config.output_path,
                        encoding=config.csv_encoding,
                        delimiter=config.csv_delimiter,
                        excel_safe_cells=True,
                    )
                    sink = QueueLogSink(self.events)
                    engine = ExtractionEngine(config, backend, writer, stats, dashboard=sink)
                    self.events.put(("stats_ref", stats, index, len(paths), pst_path))
                    # Store only for cancellation; the worker never touches Tk widgets.
                    self.active_engine = engine
                    if self.cancel_event.is_set():
                        engine.request_stop()
                    stats = engine.run()
                except Exception as exc:
                    logging.exception("Failed to process PST %s", pst_path)
                    self.events.put(("log", f"Could not process {pst_path.name}: {type(exc).__name__}: {exc}", "error"))
                    results.append({"path": pst_path, "output": output_path, "stats": stats, "failed": True})
                    self.events.put(("file_complete", pst_path, output_path, stats, True))
                    continue
                finally:
                    self.active_engine = None
                    if writer is not None:
                        try:
                            writer.close()
                        except Exception as exc:
                            logging.exception("Failed to close CSV writer")
                            self.events.put(("log", f"Output finalisation error: {exc}", "error"))

                failed = stats.phase == "failed"
                results.append({"path": pst_path, "output": output_path, "stats": stats, "failed": failed})
                if failed:
                    self.events.put(("log", f"Failed: {pst_path.name}. See {config.log_path}", "error"))
                else:
                    self.events.put(("log", f"Finished {pst_path.name}: {stats.exported:,} exported, {stats.errors:,} errors.", "ok"))
                    self.events.put(("log", f"CSV: {output_path}", "ok"))
                self.events.put(("file_complete", pst_path, output_path, stats, failed))
        except Exception as exc:
            logging.exception("Unexpected GUI worker failure")
            self.events.put(("log", f"Unexpected failure: {type(exc).__name__}: {exc}", "error"))
        finally:
            self.events.put(("job_done", results, self.cancel_event.is_set()))

    def _poll_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == "log":
                    _, message, level = event
                    self._append_log(message, level)
                elif kind == "file_start":
                    _, index, total, path = event
                    self.active_index = index
                    self.active_stats = None
                    self.current_var.set(f"File {index} of {total}: {path.name}")
                elif kind == "stats_ref":
                    _, stats, index, total, path = event
                    self.active_stats = stats
                    self.active_index = index
                    self.total_files = total
                    self.current_var.set(f"Opening {path.name}…")
                elif kind == "repair_request":
                    _, path, acknowledged, outcome = event
                    self._handle_repair_request(path, acknowledged, outcome)
                elif kind == "file_complete":
                    _, path, output_path, stats, failed = event
                    self.completed_files += 1
                    if output_path:
                        self.current_var.set(f"Completed: {path.name} → {output_path.name}")
                elif kind == "job_done":
                    _, results, cancelled = event
                    self._finish_job(results, cancelled)
        except queue.Empty:
            pass
        except Exception:
            logging.exception("GUI event queue handling failed")
        finally:
            if self.winfo_exists():
                self.after(120, self._poll_events)

    def _handle_repair_request(self, path: Path, acknowledged: threading.Event, outcome: dict[str, Any]) -> None:
        try:
            exe = find_scanpst()
            if exe is None:
                proceed = messagebox.askyesno(
                    "SCANPST not found",
                    "Microsoft's SCANPST.EXE was not found.\n\nContinue extraction without repair?",
                    parent=self,
                )
                outcome["continue"] = proceed
                if not proceed:
                    self.cancel_event.set()
                return
            launched, detail = launch_scanpst(path, wait=False)
            if not launched:
                proceed = messagebox.askyesno(
                    "Could not start repair tool",
                    f"{detail}\n\nContinue extraction without repair?",
                    parent=self,
                )
                outcome["continue"] = proceed
                if not proceed:
                    self.cancel_event.set()
                return
            self._append_log(f"Launched {exe} for {path.name}.", "warn")
            messagebox.showinfo(
                "Repair step",
                "SCANPST has been opened. Complete the scan/repair in that window, close it when finished, then click OK here to continue.\n\nIf the PST path was not prefilled, browse to:\n" + str(path),
                parent=self,
            )
            outcome["continue"] = True
        except Exception as exc:
            logging.exception("Repair tool request failed")
            self._append_log(f"Repair step error: {exc}", "error")
            outcome["continue"] = messagebox.askyesno("Repair error", "Continue extraction without repair?", parent=self)
            if not outcome["continue"]:
                self.cancel_event.set()
        finally:
            acknowledged.set()

    def _refresh_progress(self) -> None:
        if not self.winfo_exists():
            return
        if self.running:
            stats = self.active_stats
            if stats is not None:
                if stats.phase == "indexing":
                    self.status_var.set(f"Indexing folders — {stats.folders_discovered:,} found")
                elif stats.phase == "opening":
                    self.status_var.set("Opening PST…")
                elif stats.phase == "extracting":
                    self.status_var.set("Extracting messages…")
                elif stats.phase == "failed":
                    self.status_var.set("Extraction failed")
                elif stats.phase == "done":
                    self.status_var.set("Finishing current PST…")
                progress = stats.completion if stats.messages_total else 0.0
                overall = ((max(0, self.active_index - 1) + progress) / max(1, self.total_files)) * 100
                self.progress_var.set(min(100.0, overall))
                self.percent_var.set(f"{self.progress_var.get():.0f}%")
                self.metrics_var.set(
                    f"Messages: {stats.messages_seen:,} / {stats.messages_total:,}    "
                    f"Exported: {stats.exported:,}    Duplicates: {stats.duplicates:,}    "
                    f"Spam filtered: {stats.spam:,}    Errors: {stats.errors:,}    "
                    f"Elapsed: {human_duration(stats.elapsed)}    Output: {human_size(stats.bytes_written)}"
                )
            else:
                overall = (self.completed_files / max(1, self.total_files)) * 100
                self.progress_var.set(overall)
                self.percent_var.set(f"{overall:.0f}%")
                if self.cancel_event.is_set():
                    self.status_var.set("Cancelling…")
                else:
                    self.status_var.set("Preparing PST…")
        if self.winfo_exists():
            self.after(250, self._refresh_progress)

    def _finish_job(self, results: list[dict[str, Any]], cancelled: bool) -> None:
        self.running = False
        self.active_engine = None
        self.active_stats = None
        self._set_controls_running(False)
        success = sum(1 for item in results if not item.get("failed") and item.get("stats") is not None)
        failures = sum(1 for item in results if item.get("failed"))
        self.progress_var.set(100 if not cancelled and self.total_files > 0 and len(results) >= self.total_files else self.progress_var.get())
        self.percent_var.set(f"{self.progress_var.get():.0f}%")
        if cancelled:
            self.status_var.set("Cancelled")
            self._append_log("Job cancelled. Any completed CSV files have been kept.", "warn")
        elif failures:
            self.status_var.set(f"Finished with {failures} issue(s)")
        else:
            self.status_var.set("Completed successfully" if results else "No files processed")
        self.current_var.set(f"Completed {len(results)} of {self.total_files} PST file(s).")
        if results:
            exported = sum(item["stats"].exported for item in results if item.get("stats") is not None)
            self._append_log(f"Job summary: {success} completed, {failures} failed/skipped, {exported:,} messages exported.", "ok" if not failures else "warn")
            if not cancelled:
                messagebox.showinfo(
                    "PST export finished",
                    f"Completed: {success}\nFailed/skipped: {failures}\nMessages exported: {exported:,}\n\nOutput folder:\n{self.output_var.get()}",
                    parent=self,
                )
        self.worker = None

    def _append_log(self, message: str, level: str = "info") -> None:
        stamp = time.strftime("%H:%M:%S")
        tag = level if level in {"info", "ok", "warn", "error", "debug"} else "info"
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{stamp}] {message}\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _set_controls_running(self, is_running: bool) -> None:
        locked = "disabled" if is_running else "normal"
        self.start_button.configure(state=locked)
        self.cancel_button.configure(state="normal" if is_running else "disabled")
        self.open_output_button.configure(state=locked)
        self.source_entry.configure(state=locked)
        self.output_entry.configure(state=locked)
        self.browse_file_button.configure(state=locked)
        self.browse_folder_button.configure(state=locked)
        self.browse_output_button.configure(state=locked)
        self.repair_check.configure(state=locked)
        self.recursive_check.configure(state=locked)
        self.dedupe_check.configure(state=locked)
        self.spam_check.configure(state=locked)
        self.backend_combo.configure(state="disabled" if is_running else "readonly")

    def _cancel(self) -> None:
        if not self.running:
            return
        self.cancel_event.set()
        engine = self.active_engine
        if engine is not None:
            engine.request_stop()
        self.status_var.set("Cancelling… finishing the current message")
        self._append_log("Cancellation requested. The app will keep any completed output.", "warn")

    def _open_output(self) -> None:
        folder = Path(self.output_var.get().strip().strip('"')).expanduser()
        try:
            folder.mkdir(parents=True, exist_ok=True)
            import os
            import subprocess
            import sys
            if sys.platform == "win32":
                os.startfile(str(folder))  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except Exception as exc:
            messagebox.showerror("Cannot open folder", str(exc), parent=self)

    def _on_close(self) -> None:
        if self.running:
            if not messagebox.askyesno("Extraction is running", "Cancel extraction and close the app?", parent=self):
                return
            self._cancel()
            self.after(300, self._wait_before_close)
        else:
            self.destroy()

    def _wait_before_close(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            self.after(300, self._wait_before_close)
        else:
            self.destroy()


def main() -> None:
    app = PSTExportGUI()
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        logging.exception("Fatal GUI exception")
        raise
