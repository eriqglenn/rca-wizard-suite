"""Central path resolution for the RCA Wizard Suite.

All paths are resolved from this file's location, never from the current
working directory, so the suite behaves identically whether it is launched
from ``Run Wizard.bat``, a shortcut, or a terminal in any folder.

Layout (relative to ``Wizard_Files``)::

    templates/<wizard_key>/      official City templates, one folder per wizard
    sample_input/<wizard_key>/   sample workbooks/maps used for demos & QA
    core/hcad_downloader.py      single shared HCAD downloader (all wizards)

Wizard keys: ``mls``, ``mbl``, ``mls_renewal``, ``mbl_renewal``.

When the City revises a template, replace the file inside the matching
``templates/<wizard_key>`` folder; no code change is required.  Every template
also keeps a Browse button in its wizard UI as a per-case override.
"""

from __future__ import annotations

import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    # Packaged build (PyInstaller, see build/): templates, sample input and
    # docs live in an editable Wizard_Files folder next to the executable.
    WIZARD_FILES_DIR: Path = Path(sys.executable).resolve().parent / "Wizard_Files"
else:
    WIZARD_FILES_DIR: Path = Path(__file__).resolve().parents[1]
TEMPLATES_ROOT: Path = WIZARD_FILES_DIR / "templates"
SAMPLE_INPUT_ROOT: Path = WIZARD_FILES_DIR / "sample_input"

#: The one shared HCAD downloader script.  The four wizards previously shipped
#: byte-identical private copies; the suite keeps a single copy here so a fix
#: lands once and reaches every wizard (master brief: changes must propagate
#: consistently across the whole suite).
HCAD_DOWNLOADER: Path = Path(__file__).resolve().parent / "hcad_downloader.py"


def templates_dir(wizard_key: str) -> Path:
    """Return the official-templates folder for one wizard."""
    return TEMPLATES_ROOT / wizard_key


def sample_input_dir(wizard_key: str) -> Path:
    """Return the bundled sample-input folder for one wizard."""
    return SAMPLE_INPUT_ROOT / wizard_key
