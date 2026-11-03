"""Calculation-sheet downloads and fresh workbook reads for all four wizards."""

from __future__ import annotations

import hashlib
import os
import shutil
import tempfile
import time
from pathlib import Path

from .paths import sample_input_dir


def _windows_shared_bytes(path: Path) -> bytes:
    """Allow Excel/OneDrive's read, write and replacement handles to coexist.

    Python's normal Windows file open does not share DELETE access. Excel and
    sync clients can hold that access even when they allow other readers.
    This requests read access only; it does not change permissions or locks.
    """
    import ctypes
    from ctypes import wintypes
    import msvcrt

    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
        ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    kernel.CreateFileW.restype = wintypes.HANDLE
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    handle = kernel.CreateFileW(str(path), 0x80000000, 7, None, 3, 0x80, None)
    if handle == ctypes.c_void_p(-1).value:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        fd = msvcrt.open_osfhandle(handle, os.O_RDONLY | os.O_BINARY)
    except Exception:
        kernel.CloseHandle(handle)
        raise
    # The CRT descriptor now owns the Windows handle, including on read failure.
    with os.fdopen(fd, 'rb') as stream:
        return stream.read()


def read_shared_bytes(path: str | Path) -> bytes:
    """Read the saved file, allowing Windows save/replace sharing when needed."""
    source = Path(path).expanduser().resolve()
    last_error = None
    for attempt in range(4):
        try:
            try:
                return source.read_bytes()
            except PermissionError:
                if os.name != 'nt': raise
                return _windows_shared_bytes(source)
        except PermissionError as error:
            last_error = error
        except OSError as error:
            if getattr(error, 'winerror', None) not in (5, 32, 33): raise
            last_error = error
        if attempt < 3: time.sleep(.15 * (attempt + 1))
    raise PermissionError(
        f'Cannot read the saved workbook "{source.name}" yet. Finish saving in Excel, '
        'then click Refresh parcel data. If it is still locked, close that workbook '
        'or save a local copy and select it in the wizard. No older workbook was used.'
    ) from last_error


CALCULATION_TEMPLATES = {
    "mls": "MLS_Calculation_Sheet_Case_Info_v13.xlsx",
    "mbl": "MBL_Calculation_Sheet_Case_Info_v3.xlsx",
    "mls_renewal": "MLS_Renewal_Calculation_Sheet_Case_Info_v2.xlsx",
    "mbl_renewal": "MBL_Renewal_Calculation_Sheet_Case_Info_v2.xlsx",
}


def calculation_template_path(wizard_key: str) -> Path:
    """The current instructional workbook, including the City's examples."""
    return sample_input_dir(wizard_key) / CALCULATION_TEMPLATES[wizard_key]


def is_bundled_input(path: str | Path) -> bool:
    """Recognize sample inputs from this or an older installed suite."""
    p = Path(path).expanduser().resolve()
    return (p.parent.name in CALCULATION_TEMPLATES
            and p.parent.parent.name.lower() == "sample_input")


def save_calculation_sheet(wizard_key: str, destination: str | Path) -> Path:
    """Copy the workbook unchanged; never overwrite a bundled template."""
    source = calculation_template_path(wizard_key).resolve()
    destination = Path(destination).expanduser().resolve()
    if destination.suffix.lower() != ".xlsx":
        raise ValueError("Save the calculation sheet as an Excel .xlsx file.")
    if destination == source or is_bundled_input(destination):
        raise ValueError("Choose your case folder to save a working copy. "
                         "The bundled calculation templates must stay unchanged.")
    shutil.copyfile(source, destination)
    return destination


class WorkbookReadCache:
    """Reuse a recalculated copy only while the saved source is unchanged.

    Compare file contents, not just the path or timestamp: Excel and synced
    folders can replace a file at the same path. Reading never edits the
    planner's workbook. Failed reads cannot fall back to an older cached copy.
    """

    def __init__(self):
        self._key = None
        self._resolved: Path | None = None
        self._snapshot_dir = None

    def resolve(self, path: str | Path) -> str:
        from .recalc import needs_recalc, recalculated_copy

        if not str(path).strip():
            raise ValueError("Select a calculation workbook first.")
        source = Path(path).expanduser().resolve()
        for _attempt in range(2):
            try:
                locked = False
                try:
                    contents = source.read_bytes()
                except PermissionError:
                    contents = read_shared_bytes(source)
                    locked = True
                key = (str(source), hashlib.sha256(contents).digest())
            except OSError:
                self._key, self._resolved = None, None
                raise
            if (key == self._key and self._resolved and self._resolved.is_file()
                    and not (locked and self._resolved == source)):
                return str(self._resolved)
            self._key, self._resolved = None, None
            readable = source
            if locked:
                # Office/ZIP readers also use ordinary file opens. Give those
                # readers a private copy and keep it alive for this cache's life.
                if self._snapshot_dir is None:
                    self._snapshot_dir = tempfile.TemporaryDirectory(prefix='rca_workbook_read_')
                readable = Path(self._snapshot_dir.name) / (key[1].hex() + source.suffix)
                readable.write_bytes(contents)
            resolved = Path(recalculated_copy(readable)) if needs_recalc(readable) else readable
            # Do not associate a temporary copy with a newer save made while
            # Office was recalculating the previous version.
            if hashlib.sha256(read_shared_bytes(source)).digest() != key[1]:
                continue
            self._key, self._resolved = key, resolved
            return str(resolved)
        raise RuntimeError("The calculation workbook changed during refresh. "
                           "Finish saving it in Excel, then refresh again.")
