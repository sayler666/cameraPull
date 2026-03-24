#!/usr/bin/env python3
"""
Camera Pull - Fetch images from Fujifilm X-S10 via USB.

Supports two camera USB modes:
  - MTP (default): camera shows as 'Digital Camera' in Windows — no extra setup
  - Mass Storage: camera appears as a removable drive

Usage:
  uv run camera_pull.py [destination_folder]

If no destination folder is given, DEST_FOLDER below is used.
"""

import sys
import ctypes
import string
import shutil
import subprocess
from collections import defaultdict
from pathlib import Path

import questionary
from rich.console import Console
from rich.table import Table
from rich.progress import (
    BarColumn, FileSizeColumn, MofNCompleteColumn, Progress,
    TaskProgressColumn, TextColumn, TimeRemainingColumn,
    TotalFileSizeColumn, TransferSpeedColumn,
)

# ── Configure ─────────────────────────────────────────────────────────────────
DEST_FOLDER = Path(r"D:\camera")
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".raf", ".raw", ".dng", ".mov", ".mp4", ".avi"}
# ─────────────────────────────────────────────────────────────────────────────

console = Console()

_QSTYLE = questionary.Style([
    ("qmark",       "fg:cyan bold"),
    ("question",    "bold"),
    ("pointer",     "fg:cyan bold"),
    ("highlighted", "fg:cyan"),
    ("selected",    "fg:green"),
    ("instruction", "fg:ansidarkgray"),
])


def _fmt_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024 or unit == "TB":
            return f"{n:.1f} {unit}"
        n /= 1024


def _show_summary_and_select(files: list, dest: Path) -> set:
    """Rich table summary + questionary checkbox selection. Returns chosen extensions."""
    groups: dict = defaultdict(lambda: [0, 0])
    for _, _, ext, size in files:
        groups[ext][0] += 1
        groups[ext][1] += size

    already = {
        ext: sum(1 for _, fname, e, _ in files if e == ext and (dest / fname).exists())
        for ext in groups
    }

    table = Table(
        show_header=True, header_style="bold dim", box=None,
        padding=(0, 3, 0, 0), show_edge=False,
    )
    table.add_column("Type",  style="cyan", min_width=8)
    table.add_column("Files", justify="right", min_width=6)
    table.add_column("Size",  justify="right", min_width=12)
    table.add_column("",      justify="left")

    for ext in sorted(groups):
        count, total = groups[ext]
        new = count - already[ext]
        if already[ext]:
            note = f"[dim]{already[ext]} exist[/dim]  [green]{new} new[/green]"
        else:
            note = f"[green]{new} new[/green]"
        table.add_row(ext, str(count), _fmt_size(total), note)

    table.add_section()
    total_all = sum(s for _, _, _, s in files)
    table.add_row(
        "[bold]Total[/bold]",
        f"[bold]{len(files)}[/bold]",
        f"[bold]{_fmt_size(total_all)}[/bold]",
        "",
    )

    console.print()
    console.print(table)
    console.print()

    choices = []
    for ext in sorted(groups):
        new_count = groups[ext][0] - already[ext]
        if new_count == 0:
            continue
        label = f"{ext}   {new_count} new  ·  {_fmt_size(groups[ext][1])}"
        choices.append(questionary.Choice(
            title=label,
            value=ext,
            checked=(ext in {".jpg", ".jpeg"}),
        ))

    if not choices:
        console.print("[dim]Nothing new to copy.[/dim]")
        return set()

    selected = questionary.checkbox(
        "Select file types:",
        choices=choices,
        style=_QSTYLE,
    ).ask()

    console.print()
    return set(selected) if selected else set()


def _make_progress(use_bytes: bool) -> Progress:
    cols = [TextColumn("[cyan]{task.description:<38}[/cyan]"), BarColumn(bar_width=None)]
    if use_bytes:
        cols += [TaskProgressColumn(), "·", FileSizeColumn(), "/",
                 TotalFileSizeColumn(), "·", TransferSpeedColumn()]
    else:
        cols.append(MofNCompleteColumn())
    cols += ["·", TimeRemainingColumn()]
    return Progress(*cols, console=console, expand=True)


# ══════════════════════════════════════════════════════════════════════════════
# Mode 1 — Mass Storage
# ══════════════════════════════════════════════════════════════════════════════

def find_camera_drive():
    """Return (label, path) for the first removable drive with a DCIM folder, or (None, None)."""
    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    for letter in string.ascii_uppercase:
        if bitmask & 1:
            path = Path(f"{letter}:\\")
            if ctypes.windll.kernel32.GetDriveTypeW(str(path)) == 2:
                if (path / "DCIM").is_dir():
                    buf = ctypes.create_unicode_buffer(256)
                    ctypes.windll.kernel32.GetVolumeInformationW(
                        str(path), buf, 256, None, None, None, None, 0)
                    label = buf.value or path.drive
                    return label, path
        bitmask >>= 1
    return None, None


def copy_from_drive(camera_drive: Path, dest: Path):
    with console.status("[cyan]Scanning camera (Mass Storage)…[/cyan]"):
        files = []
        for src in camera_drive.rglob("*"):
            ext = src.suffix.lower()
            if ext in IMAGE_EXTENSIONS:
                files.append((src, src.name, ext, src.stat().st_size))

    if not files:
        console.print("[yellow]No media files found on camera.[/yellow]")
        return

    dest.mkdir(parents=True, exist_ok=True)
    selected_exts = _show_summary_and_select(files, dest)
    if not selected_exts:
        console.print("[dim]Nothing selected.[/dim]")
        return

    to_copy = [
        (src, fname, size)
        for src, fname, ext, size in files
        if ext in selected_exts and not (dest / fname).exists()
    ]
    skipped = sum(
        1 for _, fname, ext, _ in files
        if ext in selected_exts and (dest / fname).exists()
    )

    if not to_copy:
        console.print(f"[dim]All selected files already copied. ({skipped} skipped)[/dim]")
        return

    total_bytes = sum(s for _, _, s in to_copy)
    copied = errors = 0

    with _make_progress(use_bytes=True) as progress:
        task = progress.add_task("", total=total_bytes)
        for src, fname, size in to_copy:
            progress.update(task, description=fname)
            try:
                shutil.copy2(src, dest / fname)
                copied += 1
            except Exception as e:
                console.print(f"  [red]ERROR[/red] {fname}: {e}")
                errors += 1
            progress.advance(task, size)

    _print_summary(copied, skipped, errors)


# ══════════════════════════════════════════════════════════════════════════════
# Mode 2 — MTP via Windows Portable Device (WPD) API
# ══════════════════════════════════════════════════════════════════════════════

_PS_FIND_CAMERA = """
$shell = New-Object -ComObject Shell.Application
foreach ($item in $shell.NameSpace(17).Items()) {
    if ($item.Type -eq "Digital Camera") { Write-Output "$($item.Name)|$($item.Path)" }
}
"""


def find_camera_mtp():
    """Use Windows Shell to find 'Digital Camera' devices.
    Returns (name, pnp_id) or (None, None)."""
    result = subprocess.run(
        ["powershell.exe", "-NoProfile", "-Command", _PS_FIND_CAMERA],
        capture_output=True, text=True,
    )
    for line in result.stdout.splitlines():
        line = line.strip()
        idx = line.find("\\?\\")
        if idx > 0:
            name, _, _ = line.partition("|")
            return name.strip(), line[idx - 1:]
    return None, None


def copy_via_mtp(pnp_id: str, dest: Path) -> bool:
    """Scan, preview, then download images from a WPD/MTP camera to dest."""
    try:
        import comtypes
        import comtypes.client as cc
    except ImportError:
        console.print("[red]comtypes not installed.[/red]  Run:  uv add comtypes")
        return False

    comtypes.CoInitialize()
    pdapi = cc.GetModule("portabledeviceapi.dll")
    import comtypes.gen.PortableDeviceApiLib as pdgen

    CLSID_PortableDeviceFTM    = comtypes.GUID("{F7C0039A-4762-488A-B4B3-760EF9A1BA9B}")
    CLSID_PortableDeviceValues = comtypes.GUID("{0C15D503-D017-47CE-9016-7B3F978721CC}")
    STGM_READ = 0

    def make_key(fmtid_str, pid):
        key = pdgen._tagpropertykey()
        key.fmtid = comtypes.GUID(fmtid_str)
        key.pid = pid
        return key

    WPD_OBJECT_FILE_NAME = make_key("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C}", 12)
    WPD_OBJECT_SIZE      = make_key("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C}", 11)
    WPD_RESOURCE_DEFAULT = make_key("{E81E79BE-34F0-41BF-B53F-F1A06AE87842}", 0)

    # IStream vtable helpers — comtypes maps Read as RemoteRead (MIDL artefact)
    # which segfaults for in-process streams, so we drive the vtable directly.
    # AddRef/Release are also called manually to guarantee the MTP transaction
    # closes before the next GetStream (comtypes may delay Release via caching).
    _UlongFn = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
    _ReadFn  = ctypes.WINFUNCTYPE(
        ctypes.HRESULT,
        ctypes.c_void_p, ctypes.c_void_p, ctypes.c_ulong,
        ctypes.POINTER(ctypes.c_ulong),
    )

    def _vtbl(raw: int):
        return ctypes.cast(
            ctypes.cast(raw, ctypes.POINTER(ctypes.c_void_p))[0],
            ctypes.POINTER(ctypes.c_void_p),
        )

    def _addref(raw: int) -> None:
        _UlongFn(_vtbl(raw)[1])(raw)

    def _release(raw: int) -> None:
        _UlongFn(_vtbl(raw)[2])(raw)

    def _stream_read(raw: int, chunk_size: int) -> bytes:
        buf = ctypes.create_string_buffer(chunk_size)
        n   = ctypes.c_ulong(0)
        _ReadFn(_vtbl(raw)[3])(raw, buf, chunk_size, ctypes.byref(n))
        return bytes(buf.raw[: n.value])

    try:
        client_info = cc.CreateObject(CLSID_PortableDeviceValues,
                                      interface=pdapi.IPortableDeviceValues)
        device = cc.CreateObject(CLSID_PortableDeviceFTM,
                                 interface=pdapi.IPortableDevice)
        device.Open(pnp_id, client_info)
    except Exception as e:
        console.print(f"[red]Could not open camera:[/red] {e}")
        return False

    content = device.Content()

    def _enum_children(parent_id):
        children = []
        try:
            en = content.EnumObjects(0, parent_id, None)
            while True:
                obj_id, fetched = en.Next(1)
                if not fetched or obj_id is None:
                    break
                children.append(obj_id)
            del en
        except Exception:
            pass
        return children

    # ── Scan ──────────────────────────────────────────────────────────────────
    scanned = []

    def _scan(parent_id):
        for obj_id in _enum_children(parent_id):
            fname = obj_id
            size  = 0
            try:
                vals  = content.Properties().GetValues(obj_id, None)
                fname = vals.GetStringValue(ctypes.byref(WPD_OBJECT_FILE_NAME)) or obj_id
                try:
                    size = vals.GetUnsignedLargeIntegerValue(ctypes.byref(WPD_OBJECT_SIZE))
                except Exception:
                    size = 0
            except Exception:
                pass
            ext = Path(fname).suffix.lower() if isinstance(fname, str) else ""
            if ext in IMAGE_EXTENSIONS:
                scanned.append((obj_id, fname, ext, size))
            _scan(obj_id)

    with console.status("[cyan]Scanning camera (MTP)…[/cyan]"):
        _scan("DEVICE")

    if not scanned:
        console.print("[yellow]No media files found on camera.[/yellow]")
        device.Close()
        return False

    # ── Preview + selection ───────────────────────────────────────────────────
    dest.mkdir(parents=True, exist_ok=True)
    selected_exts = _show_summary_and_select(scanned, dest)
    if not selected_exts:
        console.print("[dim]Nothing selected.[/dim]")
        device.Close()
        return True

    # ── Filter ────────────────────────────────────────────────────────────────
    to_copy = [
        (obj_id, fname, size)
        for obj_id, fname, ext, size in scanned
        if ext in selected_exts and not (dest / fname).exists()
    ]
    skipped = sum(
        1 for _, fname, ext, _ in scanned
        if ext in selected_exts and (dest / fname).exists()
    )

    if not to_copy:
        console.print(f"[dim]All selected files already copied. ({skipped} skipped)[/dim]")
        device.Close()
        return True

    total_bytes = sum(size for _, _, size in to_copy)
    use_bytes   = total_bytes > 0

    # Done scanning — close session; reopen per-file below.
    # Fujifilm firmware gets stuck after one GetStream per session;
    # a fresh Open/Close resets the MTP state for each file.
    content = None
    device.Close()

    # ── Copy ──────────────────────────────────────────────────────────────────
    copied = errors = 0
    with _make_progress(use_bytes) as progress:
        task = progress.add_task("", total=total_bytes if use_bytes else len(to_copy))

        for obj_id, fname, size in to_copy:
            progress.update(task, description=fname)
            dst = dest / fname
            stream_raw = 0
            try:
                ci = cc.CreateObject(CLSID_PortableDeviceValues,
                                     interface=pdapi.IPortableDeviceValues)
                device.Open(pnp_id, ci)
                resources = device.Content().Transfer()

                buf_size = ctypes.c_ulong(0)
                result   = resources.GetStream(
                    obj_id,
                    ctypes.byref(WPD_RESOURCE_DEFAULT),
                    STGM_READ,
                    ctypes.pointer(buf_size),
                )
                comptr     = result[-1]
                result     = None
                stream_raw = ctypes.cast(comptr, ctypes.c_void_p).value
                _addref(stream_raw)
                comptr     = None
                resources  = None

                chunk = max(buf_size.value, 131072)
                data  = bytearray()
                while True:
                    chunk_data = _stream_read(stream_raw, chunk)
                    if not chunk_data:
                        break
                    data.extend(chunk_data)
                    if use_bytes:
                        progress.advance(task, len(chunk_data))

                _release(stream_raw)
                stream_raw = 0

                dst.write_bytes(data)
                copied += 1
                if not use_bytes:
                    progress.advance(task, 1)

            except Exception as e:
                console.print(f"  [red]ERROR[/red] {fname}: {e}")
                errors += 1
                if not use_bytes:
                    progress.advance(task, 1)
            finally:
                if stream_raw:
                    _release(stream_raw)
                    stream_raw = 0
                try:
                    device.Close()
                except Exception:
                    pass

    _print_summary(copied, skipped, errors)
    return copied > 0 or skipped > 0


# ══════════════════════════════════════════════════════════════════════════════
# Shared helpers
# ══════════════════════════════════════════════════════════════════════════════

def _print_summary(copied: int, skipped: int, errors: int) -> None:
    err_style = "red" if errors else "dim"
    console.print(
        f"\n[green]✓[/green]  "
        f"[bold]{copied}[/bold] copied  [dim]·[/dim]  "
        f"[dim]{skipped} skipped[/dim]  [dim]·[/dim]  "
        f"[{err_style}]{errors} errors[/{err_style}]"
    )


# ══════════════════════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════════════════════

def main():
    dest = Path(sys.argv[1]) if len(sys.argv) > 1 else DEST_FOLDER

    console.print(f"\n[bold cyan]Camera Pull[/bold cyan]")
    console.print(f"[dim]Destination:[/dim] {dest}")
    console.rule(style="dim")

    # ── Mass Storage ──────────────────────────────────────────────────────────
    label, drive = find_camera_drive()
    if drive:
        console.print(f"[green]✓[/green] [bold]{label}[/bold]  [dim](Mass Storage: {drive})[/dim]")
        copy_from_drive(drive, dest)
        return

    # ── MTP via WPD ───────────────────────────────────────────────────────────
    with console.status("[cyan]Looking for camera…[/cyan]"):
        cam_name, pnp_id = find_camera_mtp()

    if pnp_id:
        console.print(f"[green]✓[/green] [bold]{cam_name}[/bold]  [dim](MTP)[/dim]")
        console.rule(style="dim")
        if copy_via_mtp(pnp_id, dest):
            return
        sys.exit(1)

    # ── Not found ─────────────────────────────────────────────────────────────
    console.print(
        "\n[yellow]Camera not detected.[/yellow] Please check:\n"
        "  1. USB cable is connected and camera is ON\n"
        "  2. On the camera: MENU > Set Up > Connection Setting > USB Mode\n"
        "     [dim]• MTP/PTP — default, works out of the box[/dim]\n"
        "     [dim]• Mass Storage — also works[/dim]\n"
        "  3. Dismiss any dialog shown on the camera screen\n"
    )
    sys.exit(1)


if __name__ == "__main__":
    main()
