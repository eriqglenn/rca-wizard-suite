"""Repair a calculation workbook with a poisoned/empty formula cache.

Symptom this fixes: Step 2 shows blank in Excel even after Ctrl+Alt+F9, and the
wizard reads no Step 2 values, even though Step 1 data and the formulas are
correct.  Cause: the file was saved by a tool that wrote every formula with an
empty cached value and no recalc flag, so Excel trusts the empty cache.

This writes a repaired copy that strips the empty cache and sets
fullCalcOnLoad, so opening it in Excel recalculates everything immediately.

Usage (from Wizard_Files):
    python tools/repair_workbook.py "C:\\path\\to\\workbook.xlsx"
    python tools/repair_workbook.py "in.xlsx" --out "in_REPAIRED.xlsx"
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from core.recalc import needs_recalc, repair_in_place_copy


def main() -> int:
    ap = argparse.ArgumentParser(description="Repair an Excel formula cache")
    ap.add_argument("workbook")
    ap.add_argument("--out", help="output path (default: <name>_REPAIRED.xlsx)")
    args = ap.parse_args()

    src = Path(args.workbook)
    if not src.exists():
        print(f"Not found: {src}")
        return 1
    stale = needs_recalc(src)
    print(f"Formula cache looks {'STALE (needs repair)' if stale else 'OK'}.")
    out = Path(args.out) if args.out else src.with_name(src.stem + "_REPAIRED.xlsx")
    res = repair_in_place_copy(src, out)
    if res:
        print(f"Wrote repaired workbook:\n  {out}\n"
              "Open it in Excel — Step 2 will recalculate on open. Save to keep it.")
        return 0
    print("Repair failed.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
