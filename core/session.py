"""Per-wizard session memory for the RCA Wizard Suite.

Remembers the planner's last file selections (workbook, map, attachments,
output folder) per wizard, so reopening a wizard does not
mean re-browsing for everything.  Stored as JSON in the user's home folder,
which is per-user and writable regardless of where the suite is installed.

Rules:
- Restoring never overwrites a value the wizard has already set unless the
  saved path still exists on disk (stale paths are skipped silently).
- Saving and loading never raise; session memory is a convenience layer and
  must not be able to break document generation.
"""

from __future__ import annotations

import json
from pathlib import Path

from .workbooks import is_bundled_input

SETTINGS_PATH = Path.home() / ".rca_wizard_suite_session.json"

#: Multi-path values (e.g. HCAD records) are joined with this separator.
MULTI_SEP = "||"


def _remember_selection(key: str, value: str) -> bool:
    # A one-off document template override lasts for this session only.
    # Reopening uses the active installed manifest (including Manage Templates
    # updates), never a remembered older template. Keep real case workbooks,
    # but start with the current instructional workbook instead of an old demo.
    if key == "template" or key.endswith("_template"):
        return False
    if key in ("excel", "workbook", "map") and is_bundled_input(value):
        return False
    return True


def _read_all() -> dict:
    try:
        if SETTINGS_PATH.exists():
            data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
    except Exception:
        pass
    return {}


def save(wizard_key: str, values: dict) -> None:
    """Persist string path values for one wizard.  Never raises."""
    try:
        clean = {k: v for k, v in values.items()
                 if isinstance(v, str) and v.strip() and _remember_selection(k, v)}
        data = _read_all()
        data[wizard_key] = clean
        SETTINGS_PATH.write_text(
            json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8"
        )
    except Exception:
        pass


def load(wizard_key: str) -> dict:
    """Return the saved values for one wizard, keeping only paths that still
    exist (multi-paths are filtered item by item).  Never raises."""
    out: dict = {}
    try:
        saved = _read_all().get(wizard_key, {})
        for k, v in saved.items():
            if not isinstance(v, str) or not v.strip():
                continue
            if not _remember_selection(k, v):
                continue
            if MULTI_SEP in v:
                kept = [p for p in v.split(MULTI_SEP) if Path(p).exists()]
                if kept:
                    out[k] = MULTI_SEP.join(kept)
            else:
                p = Path(v)
                # output folders may not exist yet; files must exist
                if p.exists() or k in ("out", "output_folder"):
                    out[k] = v
    except Exception:
        return {}
    return out


def clear(wizard_key: str | None = None) -> None:
    """Forget one wizard's session (or all sessions).  Never raises."""
    try:
        if wizard_key is None:
            if SETTINGS_PATH.exists():
                SETTINGS_PATH.unlink()
            return
        data = _read_all()
        data.pop(wizard_key, None)
        SETTINGS_PATH.write_text(
            json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8"
        )
    except Exception:
        pass
