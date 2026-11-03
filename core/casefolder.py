"""Per-case output folders for the RCA Wizard Suite.

Instead of every document landing loose in ``Documents``, the suite suggests
``Documents/RCA_Cases/<application number>/`` when a workbook is loaded, so a
case's letters, forms, HCAD records, and packet stay together.

This is a *suggestion*, not a rule: it only replaces the output folder when
the current value is still the plain Documents default or a previous
RCA_Cases folder (so loading the next case rolls the folder forward).  A
custom folder the planner chose by hand is never overridden.
"""

from __future__ import annotations

import re
from pathlib import Path

CASES_DIRNAME = "RCA_Cases"


def _sanitize(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r"[^A-Za-z0-9._ -]+", "_", name)
    name = re.sub(r"\s+", "_", name).strip("._")
    return name[:80] or "Case"


_INVALID_APP_NOS = {"", "0", "0.0", "none", "null", "n/a", "na", "-", "tbd", "?"}


def _normalize_app_no(value) -> str:
    """Clean an application-number cell value; '' if it isn't usable.

    A plausible application number contains at least one digit, is short, and
    is not a sentence (at most one internal space) — this stops the Step 2
    fallback scan from grabbing narrative text near an 'Application' label.
    """
    s = str(value).strip()
    try:  # 926.0 -> 926
        f = float(s)
        if f.is_integer():
            s = str(int(f))
    except (ValueError, OverflowError):
        pass
    if s.lower() in _INVALID_APP_NOS:
        return ""
    if not any(ch.isdigit() for ch in s):
        return ""
    if len(s) > 20 or s.count(" ") > 1:
        return ""
    return s


def read_application_number(workbook_path: str | Path) -> str:
    """Best-effort read of the application number from Case Info (key/value
    layout) or Step 2 (label followed by value).  Returns '' on any problem
    or when the cell holds a junk value like 0 or N/A."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(str(workbook_path), data_only=True, read_only=True)
        try:
            # 1) Case Info tab, key in column A / value in column B
            if "Case Info" in wb.sheetnames:
                ws = wb["Case Info"]
                for row in ws.iter_rows(min_col=1, max_col=2, max_row=200):
                    key = row[0].value
                    if key and str(key).strip().lower() in (
                        "application_number", "application number", "application no",
                        "application no.", "app_number",
                    ):
                        val = _normalize_app_no(row[1].value if len(row) > 1 else None)
                        if val:
                            return val
            # 2) Step 2 sheet: a label cell followed by the value in the same row
            for name in wb.sheetnames:
                if not name.lower().startswith("step 2"):
                    continue
                ws = wb[name]
                for row in ws.iter_rows(max_row=40, max_col=12):
                    for i, cell in enumerate(row):
                        v = cell.value
                        if v and re.match(r"(?i)^\s*application\s*(number|no\.?)?\s*:?\s*$",
                                          str(v)):
                            for nxt in row[i + 1:]:
                                val = _normalize_app_no(nxt.value)
                                if val:
                                    return val
                break
        finally:
            wb.close()
    except Exception:
        pass
    return ""


def case_folder_for(workbook_path: str | Path) -> Path:
    """Documents/RCA_Cases/<application number or workbook name>."""
    app_no = read_application_number(workbook_path)
    label = _sanitize(app_no) if app_no else _sanitize(Path(workbook_path).stem)
    return Path.home() / "Documents" / CASES_DIRNAME / label


def maybe_case_folder(current_out: str | Path | None,
                      workbook_path: str | Path) -> Path | None:
    """Return the suggested case folder if the current output folder should be
    rolled forward, else ``None`` (planner chose a custom location).

    Rolls forward when the current value is empty, the plain Documents folder,
    or any folder under Documents/RCA_Cases (i.e. a previous case).
    """
    try:
        suggestion = case_folder_for(workbook_path)
        if not current_out:
            return suggestion
        cur = Path(str(current_out)).expanduser()
        docs = Path.home() / "Documents"
        if cur == docs:
            return suggestion if suggestion != cur else None
        cases_root = docs / CASES_DIRNAME
        try:
            cur.relative_to(cases_root)
        except ValueError:
            return None  # custom folder — leave it alone
        return suggestion if suggestion != cur else None
    except Exception:
        return None
