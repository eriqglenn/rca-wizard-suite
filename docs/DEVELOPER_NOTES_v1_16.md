# Calculation-sheet access and saved-workbook refresh

September 7, 2026 — suite v1.16

The planner opens the suite before completing a calculation sheet. Each existing
application card now saves the matching current instructional workbook through a
Save As dialog, then opens the saved copy in Excel. `core/workbooks.py` maps the
four application types to explicit current file names. Copying preserves all
example entries and workbook content. A destination inside a bundled sample
folder is rejected to protect the masters. Cancelling creates no file.

All wizard defaults use this same mapping. Separate case workbooks remain in
session memory; bundled samples from older suites do not. Document-template
Browse overrides are not saved or restored. Startup resolves the current
installed templates (the active manifest on the packaged line). Existing
Manage Templates replacements remain the defaults on the packaged line.

Every wizard's `_wb()` delegates to `WorkbookReadCache.resolve()`. A SHA-256
fingerprint of the saved source is checked before reuse, so a changed file at
the same path, even with the same size/timestamp, invalidates a temporary copy.
A source changed during recalculation is retried once; a missing/unreadable
source never returns the old copy. A missing temporary copy is regenerated.
The planner must save Excel edits before refreshing. Calculation rules and
Office-conversion behavior are unchanged.

The 2026 new-MBL ordinance remains encrypted; the previous usable ordinance
continues as the installed default. Other document templates and the four
instructional workbooks are byte-identical to the previous release.

Validation commands:

- `python -B tools/test_workflow.py` — downloads, cancelled Save As, current
  defaults, older session entries, all four saved-workbook readers and home
  actions at 1280x860 and 1100x760, using invisible Tk windows and temp settings.
- `python -B tools/validate.py` — existing full document validation contract.

No change to eligibility thresholds, calculation formulas, letter wording,
approval routing or packet composition is included in this release.

Validation on September 7, 2026: all 637 existing checks passed, and all six
workflow tests passed. The 91 template/sample files match v1.15 byte for byte.

Workflow research: reviewed Houston's [program page and all four currently
linked block application packets](https://houstontx.gov/planning/Min-Lot_Size-Min_Bldg_Line.html),
plus the City's linked Chapter 42 excerpts for
[MLS](https://houstontx.gov/planning/Neighborhood/docs_pdfs/SMLS_Ord.pdf) and
[MBL](https://houstontx.gov/planning/Neighborhood/docs_pdfs/SMBL_Ord.pdf).
These are the City's posted excerpts, not an independent verification of the
entire current codification. The changes support obtaining the correct workbook
before GIS calculation and map preparation, then loading the saved case into
the appropriate wizard for document generation.
