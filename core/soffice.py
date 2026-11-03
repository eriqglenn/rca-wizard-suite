"""Shared LibreOffice / soffice resolver for the RCA Wizard Suite.

Each wizard previously carried its own copy of ``_find_soffice()``.  This is
the single shared resolver so a fix to LibreOffice path handling lands once
for the whole suite (master brief: changes must propagate consistently).
"""

from __future__ import annotations

import shutil
from pathlib import Path


def find_soffice() -> str | None:
    """Return a usable soffice/libreoffice executable path, or None."""
    for exe in ("soffice", "libreoffice"):
        found = shutil.which(exe)
        if found:
            return found
    for p in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ):
        if Path(p).exists():
            return p
    return None
