"""Non-blocking preflight checks for the RCA Wizard Suite.

Surfaces the things planners otherwise discover only after generating a bad
document: missing workbook tabs, empty owner rows, blank required Case Info
fields, or no boundary map selected.  Everything here is informational —
nothing in this module ever blocks generation.  Per-document validation
stays exactly where it already lives in each wizard (each document checks
only the files it actually uses).

Also provides a small Tk panel renderer so all four wizards show the same
checklist on their Case Info page, plus a key/value snapshot used by the
renewal wizards (whose Case Info pages previously showed no data at all —
master brief section 5 calls for a quality-control screen).
"""

from __future__ import annotations

import tkinter as tk
from pathlib import Path

OK = "ok"
WARN = "warn"

#: Case Info fields every wizard needs for the mail-out.
COMMON_FIELDS = ["Letter_Date", "Protest_Deadline", "Planner_Name",
                 "Planner_Phone", "Planner_Email", "Approval_Path"]

#: Extra fields the renewal wizards rely on (warn-level only).
RENEWAL_FIELDS = ["Original_Ordinance_No", "Original_Effective_Date",
                  "Original_Expiration_Date"]

#: Key case values worth showing in the snapshot, in display order.
SNAPSHOT_FIELDS = [
    "Application_Number", "Application_Type", "Street_Name",
    "Block_Description", "Side_Of_Street", "Cross_Streets",
    "Minimum_Lot_Size", "Minimum_Building_Line",
    "Letter_Date", "Protest_Deadline",
    "Planner_Name", "Planner_Phone", "Planner_Email",
    "Approval_Path", "Council_Deadline",
    "Original_Ordinance_No", "Original_Effective_Date",
    "Original_Expiration_Date", "Owner_Letter_Type",
    "SFR_Percentage", "Total_Lot_Count", "Lots_Meeting_Minimum",
    "Council_District", "Protest_Count", "PC_Hearing_Date",
]


def _read_case_info_raw(wb) -> dict:
    """Plain Case Info key/value read (column A key, column B value)."""
    info: dict = {}
    if "Case Info" not in wb.sheetnames:
        return info
    ws = wb["Case Info"]
    for row in ws.iter_rows(min_col=1, max_col=2, max_row=300):
        k = row[0].value
        if k:
            v = row[1].value if len(row) > 1 else None
            info[str(k).strip()] = "" if v is None else str(v).strip()
    return info


def _count_owner_rows(wb) -> int | None:
    """Best-effort owner count from the Step 1 GIS tab (rows with an owner
    name in an OWNER_MAILTO/Mailto-style column).  None if undeterminable."""
    step1 = next((n for n in wb.sheetnames if n.lower().startswith("step 1")), None)
    if not step1:
        return None
    ws = wb[step1]
    owner_col = None
    header_row = None
    for r, row in enumerate(ws.iter_rows(max_row=10), start=1):
        for c, cell in enumerate(row, start=1):
            v = str(cell.value).strip().lower() if cell.value else ""
            if v in ("owner_mailto", "mailto", "owner name", "owner_name"):
                owner_col, header_row = c, r
                break
        if owner_col:
            break
    if not owner_col:
        return None
    count = 0
    for row in ws.iter_rows(min_row=header_row + 1, min_col=owner_col,
                            max_col=owner_col):
        v = row[0].value
        if v not in (None, "") and str(v).strip():
            count += 1
    return count


def run_preflight(workbook_path: str | Path | None,
                  wizard_key: str,
                  map_path: str | Path | None = None,
                  output_folder: str | Path | None = None) -> list[tuple[str, str]]:
    """Return [(level, message)] — level is 'ok' or 'warn'.  Never raises and
    never blocks anything; this is a checklist, not a gate."""
    results: list[tuple[str, str]] = []
    try:
        if not workbook_path or not Path(str(workbook_path)).exists():
            results.append((WARN, "No calculation workbook selected yet."))
            return results

        try:
            from openpyxl import load_workbook
            wb = load_workbook(str(workbook_path), data_only=True, read_only=True)
        except Exception as e:
            results.append((WARN, f"Workbook could not be opened: {e}"))
            return results

        try:
            names = [n.lower() for n in wb.sheetnames]
            for label, prefix in [("Step 1 (GIS properties)", "step 1"),
                                  ("Step 2 (calculation sheet)", "step 2"),
                                  ("Case Info", "case info")]:
                if any(n.startswith(prefix) for n in names):
                    results.append((OK, f"{label} tab found."))
                else:
                    results.append((WARN, f"{label} tab is missing."))

            owners = _count_owner_rows(wb)
            if owners is None:
                results.append((WARN, "Could not count owner rows on Step 1 "
                                      "(owner column not recognized)."))
            elif owners == 0:
                results.append((WARN, "Step 1 has no owner rows — notification "
                                      "letters would come out empty."))
            else:
                results.append((OK, f"Step 1 has {owners} owner row(s)."))

            ci = _read_case_info_raw(wb)
            fields = list(COMMON_FIELDS)
            if "renewal" in wizard_key:
                fields += RENEWAL_FIELDS
            missing = [f for f in fields if not ci.get(f, "").strip()]
            if missing:
                results.append((WARN, "Blank Case Info field(s): "
                                      + ", ".join(missing)))
            else:
                results.append((OK, "Required Case Info fields are filled in."))
        finally:
            try:
                wb.close()
            except Exception:
                pass

        # Stale/poisoned formula cache: Step 2 shows blank in Excel and reads as
        # empty here even though the formulas and data are fine.  Flag it clearly
        # because it is the cause of "Step 2 is not transferring from Step 1".
        try:
            from core.recalc import needs_recalc
            if needs_recalc(workbook_path):
                results.append((WARN, "Step 2 formulas are not calculated in this "
                                      "file (empty formula cache). The wizard will "
                                      "recalculate it automatically when generating; "
                                      "to fix the file itself, use \u201cRepair workbook "
                                      "for Excel\u201d or open it in Excel, press "
                                      "Ctrl+Alt+F9, and Save."))
            else:
                results.append((OK, "Step 2 formulas are calculated."))
        except Exception:
            pass

        if map_path and Path(str(map_path)).exists():
            results.append((OK, "Boundary map selected."))
        else:
            results.append((WARN, "No boundary map selected — needed for "
                                  "notification letters (other documents do "
                                  "not require it)."))

        if output_folder and str(output_folder).strip():
            results.append((OK, f"Output folder: {output_folder}"))
        else:
            results.append((WARN, "No output folder selected."))
    except Exception as e:
        results.append((WARN, f"Preflight could not finish: {e}"))
    return results


# ----------------------------------------------------------------------------
# Tk rendering helpers (plain tk so they sit cleanly in every wizard's theme)
# ----------------------------------------------------------------------------

def build_preflight_panel(parent, results: list[tuple[str, str]],
                          card_bg: str = "#FFFFFF",
                          text_color: str = "#0F172A") -> tk.Frame:
    """Render the checklist as ✓ / ⚠ lines inside a card-style frame."""
    panel = tk.Frame(parent, bg=card_bg)
    if not results:
        results = [(WARN, "No preflight information available.")]
    for level, msg in results:
        icon = "\u2713" if level == OK else "\u26A0"
        color = "#15803D" if level == OK else "#B45309"
        row = tk.Frame(panel, bg=card_bg)
        row.pack(fill="x", anchor="w")
        tk.Label(row, text=icon, bg=card_bg, fg=color,
                 font=("Segoe UI", 10, "bold"), width=2).pack(side="left")
        tk.Label(row, text=msg, bg=card_bg, fg=text_color,
                 font=("Segoe UI", 9), wraplength=820,
                 justify="left").pack(side="left", anchor="w", pady=1)
    return panel


def build_snapshot_panel(parent, case_info: dict,
                         card_bg: str = "#FFFFFF",
                         label_color: str = "#334155",
                         value_color: str = "#0F172A") -> tk.Frame:
    """Render key Case Info values as a two-column grid (quality-control
    view for the Case Info page)."""
    panel = tk.Frame(parent, bg=card_bg)
    shown = [(f, case_info.get(f, "")) for f in SNAPSHOT_FIELDS
             if str(case_info.get(f, "")).strip()
             and str(case_info.get(f, "")).strip().lower() not in ("none", "null")]
    if not shown:
        tk.Label(panel, text="No Case Info values were read from the workbook yet.",
                 bg=card_bg, fg=label_color, font=("Segoe UI", 9)).pack(
            anchor="w", padx=2, pady=2)
        return panel
    for i, (field, value) in enumerate(shown):
        r, c = divmod(i, 2)
        cell = tk.Frame(panel, bg=card_bg)
        cell.grid(row=r, column=c, sticky="w", padx=(2, 24), pady=2)
        tk.Label(cell, text=field.replace("_", " ") + ":", bg=card_bg,
                 fg=label_color, font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Label(cell, text=str(value), bg=card_bg, fg=value_color,
                 font=("Segoe UI", 9), wraplength=330,
                 justify="left").pack(side="left", padx=(6, 0))
    return panel
