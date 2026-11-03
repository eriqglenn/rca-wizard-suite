from __future__ import annotations
import argparse
from pathlib import Path
import sys as _sys
_sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from wizards.mls.mailout_engine import generate_letters, summarize_workbook

parser = argparse.ArgumentParser(description="Generate owner notification mail-out letters.")
parser.add_argument("--excel", required=True, help="Calculation workbook with Step 1, Step 2, and Case Info tabs")
parser.add_argument("--template", required=True, help="Clean Word notification template")
parser.add_argument("--map", required=True, help="Map image/PDF to insert into each letter")
parser.add_argument("--output", required=True, help="Output DOCX path")
parser.add_argument("--preview", action="store_true", help="Generate only the first valid owner letter")
parser.add_argument("--validate-only", action="store_true", help="Read workbook and print summary/warnings without generating")
args = parser.parse_args()

if args.validate_only:
    summary = summarize_workbook(args.excel, args.map)
    print(f"Owners: {len(summary.owners)}")
    print("Subject:")
    print(summary.subject_line)
    if summary.warnings:
        print("Warnings:")
        for w in summary.warnings:
            print(f"- {w}")
    else:
        print("No warnings.")
else:
    count = generate_letters(Path(args.excel), Path(args.template), Path(args.map), Path(args.output), preview=args.preview)
    print(f"Generated {args.output} with {count} letter(s).")
