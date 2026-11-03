# Suite v1.15 — developer notes

Companion to `DEVELOPER_NOTES_v1_14.md`. v1.15 moves the MBL (new
application) wizard to the 2026 City templates and onto the shared
wording-based engine. All four wizards now fill the City's own files by
wording; the only token template left is the MBL draft ordinance.

## What changed

- **`wizards/mbl/wizard.py`** (v9.0) was rewritten: Step 2 reader
  (`_get_step2`, with a rebuild from the typed Step 1 / Step 2 columns when
  the formula cache is empty), context builder (`_get_context`, incl. the
  staff-report variant), generators for the owner notification letter, notice
  of public hearing (support under 51%, new), protest letter (new), notice of
  PC rescheduling (new), applicant sign letter, Director approval (2026 form,
  converted from the City's .doc), Planning Commission approval (City .doc
  converted), routing form (content controls), three OneMeeting staff reports
  (new) and the RCA packet. Merge helpers come from `core/docx_merge.py`;
  `_fill_sdts_in_order` from `wizards/mls/templates_2026.py`; `_fill_runs`
  and `_replace_line_in_paragraph` from `wizards/mbl_renewal/wizard.py`.
- **`_unwrap_smart_tags`**: the City's converted .doc approval forms wrap
  sample addresses in Word smart tags (`w:smartTag`), which python-docx does
  not expose as runs. The wizard unwraps them before filling.
- **`wizards/mbl/mbl_mailout_engine.py`** stays for the draft ordinance
  (token template), the Step 2 staff-analysis export and the HCAD address
  list only. The City's 2026 ordinance file is password-encrypted
  (`EncryptedPackage` / `StrongEncryptionDataSpace`); once an unlocked copy
  arrives the ordinance can follow the MBL Renewal pattern.
- **Templates** (`templates/mbl/`): 2026 files installed, old ones in
  `_previous/`; manifest roles `hearing_template`, `protest_template`,
  `reschedule_template`, `staff_report_template`,
  `staff_report_less51_template`, `staff_report_protest_template` added
  (`docs/mbl/TEMPLATES_2026.md`).
- **Workbook**: `sample_input/mbl/MBL_Calculation_Sheet_Case_Info_v3.xlsx`
  (+ filled sample, application 1030) built with Excel COM from the City's
  2026 sheet; the v2 workbooks were removed.
- `core/templates.py`, `templates/manifest.json`, `app.py` (MBL v9.0),
  `core/embedding.py` `SUITE_VERSION = "v1.15"`.
- `tools/validate.py`: `validate_mbl` rewritten for the new API (four owner
  letters, three staff reports, stale-cache and datetime regressions);
  `check_contains` now treats non-breaking spaces as spaces (the City letters
  set `15{nbsp}feet`).

- **`core/docx_merge.py`** (and the MLS Renewal wizard's helper copy):
  `_insert_map` now sizes the map with `_map_fit_width` (fits the width and
  the height of the last section's text area, 0.35" safety) and drops empty
  paragraphs after the map; `_append_letter_as_own_section` gives the
  section-break paragraph zero height. Word used to add a blank page after
  a letter whenever the map was taller than the page body (e.g. an 11x17
  or legal-size GIS export). Verified with Word page counts for letter,
  legal, tabloid, landscape, square and tall-PNG maps in all four wizards
  (`scratchpad/maptest.py`).

- **Field shading**: the City letter templates carry Word's grey field
  shading (`w:shd` fill E1E3E6) on the «merge field» runs of the address
  block. `_replace_in_runs` (core and the MLS Renewal copy) now calls
  `_clear_run_shading` whenever a merge-field placeholder was replaced, so
  no letter prints a grey address (`scratchpad/shadingtest.py` checks all
  18 letter types).

## Validation

`python tools/validate.py` must stay `RESULT: PASS`. The MBL run checks,
among others, `Block Application No. 1030` on the routing form, `15 feet` in
the letters and approvals, `owners of 60% of the SMBLB` and `Three (3) out of
four (4) developed properties` on the PC approval and `fifteen (15) feet` in
the ordinance.

The same change set was applied to the v1.10 code line as
`RCA_Wizard_Suite v1_10_4` (`scratchpad/port_mbl_v10.py`).

## Still open (planned consolidation)

The MLS Renewal wizard still carries its own copy of the merge helpers and
the MBL ordinance is the last token template. Next: one engine for all four
wizards, placeholder maps as data, a uniform leftover-placeholder scan in the
wizards and the validator.
