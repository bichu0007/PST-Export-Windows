# PST Export Toolkit — Windows Desktop App

A Windows desktop application for exporting Outlook `.pst` mailboxes to Excel-friendly CSV files. The app uses a native Tkinter GUI and the same extraction engine as the command-line tool.

## One-click launch

1. Install **Python 3.10 or newer** from <https://www.python.org/downloads/windows/>. During setup, enable the option to add Python to PATH. Use a Python distribution that includes Tkinter.
2. Extract this whole ZIP folder to a normal location, such as `C:\Tools\PST-Export-Windows`.
3. Double-click **`Open PST Export.bat`**. On first launch it creates `.venv`, installs the standard packages, and attempts to install a compatible prebuilt native PST reader. It then opens the desktop window without leaving a terminal window open.
4. Double-click **`Create Desktop Shortcut.bat`** once to create a shortcut on your Windows desktop. After that, launch from the shortcut.

Initial setup needs internet access to install packages. If the native reader cannot be installed, install Microsoft Outlook and use the Outlook reader (`pywin32`) where supported.

## GUI features

- **Browse PST File…** selects one `.pst` file.
- **Browse Folder…** processes `.pst` files in that folder. `Search subfolders in folder mode` is enabled by default.
- **Output location** defaults to `extracted_output` next to the app. The folder is created automatically and can be changed in the GUI.
- **Scan/repair with SCANPST** opens Microsoft's inbox repair utility and pauses extraction while the operator completes the repair. If SCANPST is not installed, the app asks whether to continue without repair. Repair is a manual Office operation; the app does not claim to repair a PST itself.
- **Live status** shows phase, overall progress, message counters, elapsed time, current PST, and output size.
- **Activity log** shows progress and errors while work is running. Detailed application diagnostics are written to `logs\gui.log`; per-export extraction errors go to the corresponding `.repair.log`.
- **Cancel** requests a graceful stop and keeps completed CSV files.

Each PST is written to its own timestamped CSV. The CSV contains date, sender, recipients, subject, plain-text body, attachment metadata, PST folder path, message ID, and size fields. It does not export the attachment files themselves.

## UI layout

1. Header — application title and purpose.
2. Source files — path field plus file and folder pickers.
3. Output location — editable folder path and browser.
4. Processing options — repair, recursive folder discovery, duplicate and spam filters, backend selector.
5. Actions — start, cancel, and open output folder.
6. Live status — progress bar, stage, current PST, counts, elapsed duration, and estimated output size.
7. Activity log — scrolling messages tagged by severity.

## Project structure

```text
PST-Export-Windows/
├── gui.py                         # Tkinter desktop GUI and worker orchestration
├── main.py                        # Existing command-line entry point
├── Open PST Export.bat             # One-click setup and desktop GUI launcher
├── Open PST Export CLI.bat         # Optional command-line launcher
├── Create Desktop Shortcut.bat
├── Create Desktop Shortcut.ps1
├── run.bat                        # CLI compatibility launcher
├── requirements.txt
├── tests/
│   └── test_core.py
└── pstexport/
    ├── __init__.py
    ├── config.py
    ├── errors.py
    ├── extractor.py               # Two-phase folder indexing and email extraction
    ├── filters.py                 # Spam and duplicate filters
    ├── models.py                  # RunConfig, MessageRecord, Stats
    ├── repair.py                  # PST checks and SCANPST discovery/launcher
    ├── ui.py                      # Existing Rich terminal dashboard for CLI
    ├── utils.py                   # Header/body normalisation helpers
    ├── writers.py                 # Streaming, Excel-safe CSV writer
    └── backends/
        ├── __init__.py
        ├── base.py
        ├── pypff_backend.py
        └── outlook_backend.py
```

## Implementation logic

1. The GUI validates the source and output paths and creates the output directory.
2. For file mode, it queues one PST. For folder mode, it discovers `.pst` files (optionally recursively) and sorts them by path.
3. Work runs on a daemon worker thread so Tkinter remains responsive. The worker sends log and repair-request events to a thread-safe queue; the main thread updates widgets.
4. For each PST, the app optionally starts SCANPST and waits for the user to confirm the manual scan/repair step.
5. `probe_pst` performs non-destructive preflight checks; the selected backend opens the PST; `ExtractionEngine` indexes folders and streams messages to `CsvStreamWriter`.
6. Progress and counters are read from the engine's `Stats` object. Errors are logged and, where safe, the app continues to the next message or next PST.
7. The GUI displays a final summary and leaves all completed CSV output in the selected directory.

## Reader backends

- `auto`: prefer `pypff`, then Outlook COM.
- `pypff`: requires an installed compatible `libpff-python` distribution exposing the `pypff` import.
- `outlook`: requires Windows, `pywin32`, and a working Outlook installation/profile.

PST-reader wheels are not available for every Python/Windows combination. The one-click launcher only attempts a **prebuilt binary** install and does not compile native code. If neither backend is available, the GUI explains which reader is missing.

## Run the CLI (optional)

After the first setup, double-click `Open PST Export CLI.bat`, or run:

```bat
.venv\Scripts\python.exe main.py --pst "C:\Mail\archive.pst" --out "C:\Out\archive.csv" --yes
```

## Validation notes

The Python source can be syntax-compiled and the core tests can be run with:

```bat
.venv\Scripts\python.exe -m pytest -q
```

A real extraction still needs a valid PST and one of the supported reader backends. `SCANPST.EXE` is supplied by Microsoft Office; it is not bundled with this project.
