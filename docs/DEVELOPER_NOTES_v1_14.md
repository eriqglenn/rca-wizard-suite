# Suite v1.14 — developer notes

Companion to `DEVELOPER_NOTES_v1_13.md`. v1.14 moves the MBL Renewal wizard
to the 2026 City templates and onto the shared wording-based engine.

## What changed

- **`wizards/mbl_renewal/wizard.py`** (v2.0) was rewritten: Step 2 reader
  with 2026 / pre-2026 layout detection (`_step2_layout`), context builder,
  and generators for the owner letter, protest letter, notice of PC
  rescheduling (new), application acceptance letter (new), applicant sign
  letter, Director approval (2026 form, converted from the City's .doc),
  Planning Commission approval (unchanged form), routing form (content
  controls), OneMeeting staff report (new), draft renewal ordinance and the
  RCA packet. Merge helpers come from `core/docx_merge.py`; ordinance
  paragraph/exhibit helpers are imported from `wizards/mls_renewal/wizard.py`.
- **`core/docx_merge.py`**: `_fill_phone` also handles `832-393` + `-` +
  `XXX` run splits; `PlannerName@houstontx.gov` added to the e-mail
  placeholders.
- **Templates** (`templates/mbl_renewal/`): 2026 files installed, old ones in
  `_previous/`; manifest roles `protest_template`, `reschedule_template`,
  `acceptance_template`, `staff_report_template` added
  (`docs/mbl_renewal/TEMPLATES_2026.md`).
- **Workbook**: `sample_input/mbl_renewal/MBL_Renewal_Calculation_Sheet_Case_Info_v2.xlsx`
  (+ filled sample, 88REN) built with Excel COM from the City's 2026 sheet;
  the old sample workbook was removed.
- `core/templates.py`, `templates/manifest.json`, `tools/validate.py`
  (renewal validator now also exercises the MBL protest / rescheduling /
  acceptance / staff-report documents), `app.py` (MBL Renewal v2.0),
  `core/embedding.py` `SUITE_VERSION = "v1.14"`.

## Validation

`python tools/validate.py` must stay `RESULT: PASS`. The MBL Renewal run
checks, among others, `Previous ordinance: 2006-63` on the routing form,
`35 feet` on the Director approval and `thirty-five (35) feet` in the
ordinance.

The same change set was applied to the v1.10 code line as
`RCA_Wizard_Suite v1_10_3` (`scratchpad/port_mblren_v10.py`).

## Still open (planned consolidation)

The MLS Renewal wizard still carries its own copy of the merge helpers; the
MBL (new application) wizard still uses its own engine and one `{{token}}`
ordinance. Both are scheduled for the engine consolidation and the uniform
leftover-placeholder scan once the 2026 MBL template set arrives.
