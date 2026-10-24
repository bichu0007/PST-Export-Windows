"""Preflight diagnostics and SCANPST.EXE integration."""
from __future__ import annotations

import os
import shutil
import struct
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

PST_MAGIC = b"!BDN"          # 0x21 0x42 0x44 0x4E — first 4 bytes of any PST/OST
MIN_PLAUSIBLE_SIZE = 64 * 1024
MIN_FREE_BYTES = 250 * 1024 * 1024

_SCANPST_GLOBS = (
    r"Microsoft Office\root\Office*\SCANPST.EXE",
    r"Microsoft Office\Office*\SCANPST.EXE",
    r"Microsoft Office*\SCANPST.EXE",
    r"Microsoft\Office*\SCANPST.EXE",
)


@dataclass
class ProbeResult:
    path: Path
    exists: bool = False
    size: int = 0
    magic_ok: bool = False
    appears_locked: bool = False
    free_space_ok: bool = True
    fatal: Optional[str] = None
    warnings: list[str] = field(default_factory=list)

    @property
    def usable(self) -> bool:
        return self.exists and self.fatal is None


def probe_pst(path: Path, output_dir: Optional[Path] = None) -> ProbeResult:
    """Non-destructive sanity check. Never raises."""
    path = Path(path)
    result = ProbeResult(path=path)

    if not path.exists():
        result.fatal = f"File not found: {path}"
        return result
    if path.is_dir():
        result.fatal = f"Expected a file, got a directory: {path}"
        return result

    result.exists = True
    try:
        result.size = path.stat().st_size
    except OSError as exc:
        result.fatal = f"Cannot stat file: {exc}"
        return result

    if result.size < MIN_PLAUSIBLE_SIZE:
        result.warnings.append(
            f"File is only {result.size:,} bytes — unusually small for a PST."
        )

    # --- magic bytes ---------------------------------------------------
    try:
        with path.open("rb") as fh:
            header = fh.read(4)
        if header == PST_MAGIC:
            result.magic_ok = True
        else:
            result.warnings.append(
                f"Header is {header!r}, expected {PST_MAGIC!r}. "
                "File may be a renamed OST, truncated, or corrupt."
            )
    except OSError as exc:
        result.fatal = f"Cannot read file header: {exc}"
        return result

    # --- lock detection -------------------------------------------------
    # Outlook holds a shared read lock; a share-deny-none reader still works.
    # A *failed* read open means something has it exclusively (rare) or the
    # volume dropped.
    try:
        with path.open("rb") as fh:
            fh.read(1)
    except OSError:
        result.appears_locked = True
        result.warnings.append("File could not be opened for reading — it may be locked.")

    for lock_name in (f"~${path.stem}.pst", f"{path.name}.lock"):
        lock_file = path.parent / lock_name
        if lock_file.exists():
            result.warnings.append(f"Lock file present: {lock_file.name}")
            break

    # --- free space -----------------------------------------------------
    target = Path(output_dir) if output_dir else path.parent
    try:
        usage = shutil.disk_usage(target if target.exists() else target.parent)
        # A CSV of a mailbox is typically a small fraction of the PST, but be
        # conservative: expect up to ~2x the PST size in free space.
        needed = min(max(result.size * 2, MIN_FREE_BYTES), 50 * 1024**3)
        if usage.free < needed:
            result.free_space_ok = False
            result.warnings.append(
                f"Low free space on {target.drive or target}: "
                f"{shutil.disk_usage(target).free / 1024**3:.1f} GB free, "
                f"~{needed / 1024**3:.1f} GB suggested."
            )
    except OSError:
        pass

    return result


def find_scanpst() -> Optional[Path]:
    """Locate SCANPST.EXE across common Office install layouts."""
    candidates: list[Path] = []

    for root_var in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
        root = os.environ.get(root_var)
        if not root:
            continue
        base = Path(root)
        if not base.is_dir():
            continue
        for pattern in _SCANPST_GLOBS:
            candidates.extend(base.glob(pattern))
        # Deep search as a last resort for Click-to-Run installs.
        office = base / "Microsoft Office"
        if office.is_dir():
            candidates.extend(office.glob("**/SCANPST.EXE"))

    if not candidates:
        which = shutil.which("SCANPST.EXE") or shutil.which("scanpst")
        if which:
            return Path(which)
        return None

    # Prefer the newest, highest-versioned build.
    candidates.sort(key=lambda p: (len(p.parts), str(p)), reverse=True)
    return candidates[0]


def launch_scanpst(pst_path: Path, wait: bool = False) -> tuple[bool, str]:
    """Start the Outlook inbox repair tool. Returns (launched, message)."""
    if sys.platform != "win32":
        return False, "SCANPST.EXE is Windows-only."

    exe = find_scanpst()
    if exe is None:
        return False, "SCANPST.EXE not found. Install Microsoft Office or Office Repair Tools."

    try:
        creation = 0
        if sys.platform == "win32":
            creation = subprocess.CREATE_NEW_PROCESS_GROUP  # type: ignore[attr-defined]
        proc = subprocess.Popen([str(exe), str(pst_path)], creationflags=creation)
        if wait:
            proc.wait()
        return True, f"Launched {exe.name} for {pst_path.name}"
    except OSError as exc:
        return False, f"Failed to launch SCANPST.EXE: {exc}"
