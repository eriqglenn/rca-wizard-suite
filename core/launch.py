"""How the suite launches its helper programs, in source and frozen builds.

The wizards run two helpers as separate processes:

  * the HCAD downloader (``core/hcad_downloader.py``)
  * the validation harness (``tools/validate.py``), used by the Template
    Manager to check a replaced template

When the suite runs from source those are plain ``python script.py`` calls.
In the packaged build (PyInstaller, see ``build/``) there is no Python
interpreter and no ``.py`` files; the same helpers live inside the console
companion executable ``RCA Wizard Tools.exe`` and are selected by a
sub-command (``hcad`` / ``validate`` / ``onboard``).

Everything that spawns a helper goes through this module so the two layouts
stay in sync.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from .paths import HCAD_DOWNLOADER, WIZARD_FILES_DIR

FROZEN: bool = bool(getattr(sys, "frozen", False))
TOOLS_EXE_NAME = "RCA Wizard Tools.exe"


def app_dir() -> Path:
    """Folder holding the running program (exe folder when frozen)."""
    if FROZEN:
        return Path(sys.executable).resolve().parent
    return WIZARD_FILES_DIR


def tools_exe() -> Path:
    return app_dir() / TOOLS_EXE_NAME


def hcad_available() -> bool:
    return tools_exe().exists() if FROZEN else HCAD_DOWNLOADER.exists()


def hcad_command() -> list[str]:
    """Command prefix that runs the HCAD downloader; append its arguments."""
    if FROZEN:
        return [str(tools_exe()), "hcad"]
    return [sys.executable, str(HCAD_DOWNLOADER)]


def hcad_cwd() -> str:
    return str(app_dir() if FROZEN else HCAD_DOWNLOADER.parent)


def validate_command(wizard: str | None = None, keep_dir: str | Path | None = None) -> list[str]:
    """Command that runs the validation harness headlessly."""
    if FROZEN:
        cmd = [str(tools_exe()), "validate"]
    else:
        cmd = [sys.executable, str(WIZARD_FILES_DIR / "tools" / "validate.py")]
    if wizard:
        cmd += ["--wizard", wizard]
    if keep_dir:
        cmd += ["--out", str(keep_dir)]
    return cmd


def onboard_command(wizard: str | None = None) -> list[str]:
    if FROZEN:
        cmd = [str(tools_exe()), "onboard"]
    else:
        cmd = [sys.executable, str(WIZARD_FILES_DIR / "tools" / "onboard_templates.py")]
    if wizard:
        cmd += ["--wizard", wizard]
    return cmd


def console_flags() -> int:
    """creationflags for a helper the planner should SEE (HCAD progress).

    From source the helper shares the launcher's console window.  The packaged
    GUI has no console, so give the helper its own window."""
    if sys.platform == "win32" and FROZEN:
        return subprocess.CREATE_NEW_CONSOLE
    return 0


def hidden_flags() -> int:
    """creationflags for a helper whose output is captured (no window)."""
    if sys.platform == "win32":
        return subprocess.CREATE_NO_WINDOW
    return 0
