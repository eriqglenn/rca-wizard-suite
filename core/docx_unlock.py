"""Make generated Word documents fully editable.

Several City templates (ordinances, routing forms) ship with Word document
protection turned on: ``<w:documentProtection w:edit="forms" w:enforcement="1">``.
Word then locks the whole document so a planner can only type into form fields
and cannot edit the surrounding text. The wizards copy the template faithfully,
so every generated document inherits that lock.

Draft ordinances are meant to be edited by the planner and handed to Legal, and
the routing-form checkboxes are written directly into the OOXML (both the
``w14:checked`` state and the visible ☒/☐ glyph) BEFORE this unlock runs, so
removing the lock never changes how a checkbox displays. The rule applied by the
wizards is therefore the simplest one: every generated Word document is saved
fully editable.

``unlock_docx`` edits only ``word/settings.xml`` inside the saved .docx, removing
the protection-enforcement elements. It touches nothing else — no body content,
fonts, sections, footers, or form-field data — and never modifies the template
on disk (it runs on the wizard's OUTPUT file). It is safe to call on a document
that has no protection (it simply does nothing).
"""

from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from pathlib import Path

_SETTINGS = "word/settings.xml"

# Protection-related elements to strip from settings.xml. writeProtection locks
# the whole file read-only; documentProtection with enforcement="1" imposes the
# forms/read-only edit restriction. Removing the elements (rather than flipping
# enforcement to "0") also drops the stored password hash cleanly.
_PROTECTION_PATTERNS = (
    re.compile(r"<w:documentProtection\b[^>]*/>"),
    re.compile(r"<w:documentProtection\b[^>]*>.*?</w:documentProtection>", re.S),
    re.compile(r"<w:writeProtection\b[^>]*/>"),
    re.compile(r"<w:writeProtection\b[^>]*>.*?</w:writeProtection>", re.S),
)


def docx_is_protected(docx_path: str | Path) -> bool:
    """True if the .docx enforces document or write protection."""
    try:
        z = zipfile.ZipFile(str(docx_path))
        if _SETTINGS not in z.namelist():
            return False
        s = z.read(_SETTINGS).decode("utf-8", "replace")
        return ('w:enforcement="1"' in s and "documentProtection" in s) \
            or "writeProtection" in s
    except Exception:
        return False


def unlock_docx(docx_path: str | Path) -> bool:
    """Remove Word document/write protection from a saved .docx in place.

    Returns True if protection was found and removed, False if there was
    nothing to do or the file could not be processed. Never raises; on any
    error the original file is left untouched.
    """
    src = Path(docx_path)
    try:
        zin = zipfile.ZipFile(str(src))
        if _SETTINGS not in zin.namelist():
            return False
        settings = zin.read(_SETTINGS).decode("utf-8", "replace")
        new_settings = settings
        for pat in _PROTECTION_PATTERNS:
            new_settings = pat.sub("", new_settings)
        if new_settings == settings:
            return False  # nothing protected

        # Rewrite the archive with only settings.xml changed.
        tmp_fd = tempfile.NamedTemporaryFile(delete=False, suffix=".docx",
                                             dir=str(src.parent))
        tmp_path = Path(tmp_fd.name)
        tmp_fd.close()
        with zipfile.ZipFile(str(tmp_path), "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == _SETTINGS:
                    data = new_settings.encode("utf-8")
                zout.writestr(item, data)
        zin.close()
        shutil.move(str(tmp_path), str(src))
        return True
    except Exception:
        try:
            if "tmp_path" in dir() and tmp_path.exists():
                tmp_path.unlink()
        except Exception:
            pass
        return False
