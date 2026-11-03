# Suite v1.13 — developer notes

Companion to `DEVELOPER_NOTES_v1_12.md`. v1.13 moves the MLS (new
application) wizard to the 2026 City templates. The MBL wizards are untouched;
the MLS Renewal wizard is unchanged apart from now having a sibling copy of
its merge helpers in `core/docx_merge.py`.

## What changed

- **`core/docx_merge.py`** (new): the wording-based merge engine extracted
  from the MLS Renewal wizard — non-breaking-space aware run replacement,
  content-control fill/unwrap, checkbox controls by position, per-letter list
  numbering restart, mailto retarget, header/footer parts via relationships,
  section-preserving letter combiner, number words. The renewal wizard still
  carries its own verified copy; a later clean-up can point it at the core
  module.
- **`wizards/mls/templates_2026.py`** (new): all MLS document generators for
  the 2026 templates. Case data still comes from `mailout_engine`
  (`build_ordinance_context` + `load_case_and_owners`), so Step 2 numbers and
  the planner-entered `Lots_Meeting_Minimum` behave exactly as in v14.
  `build_context` lower-cases the side words, builds the area phrases and the
  City-style lot-size words (`ten thousand three hundred and fifty`).
- **`wizards/mls/mailout_engine.py`**: `OWNER_OPTIONAL_FIELD_MAP` reads
  `Legal_Dscr_1/2` for the owner letters. Nothing else changed; the old
  token generators remain in the module but are no longer called by the UI.
- **`wizards/mls/wizard.py`** (v15.0): template rows and buttons for the
  protest letter, the notice of public hearing (support under 51 %) and the
  OneMeeting staff report; packet requires only the tax map (plat, HCAD,
  staff analysis optional); remembered template paths that belong to another
  suite copy are ignored (`_is_foreign_bundled_template`).
- **Templates** (`templates/mls/`): 2026 files installed under wizard names
  (see `docs/mls/TEMPLATES_2026.md`), old ones in `_previous/`, manifest roles
  `protest_template`, `hearing_template`, `staff_report_template` added.
- **Workbook**: `sample_input/mls/MLS_Calculation_Sheet_Case_Info_v13.xlsx`
  (+ filled sample) built with Excel COM from the City's 2026 sheet; the v12
  workbook was removed. Read-only-recommended flag stripped.
- `core/templates.py` `DEFAULT_ROLES["mls"]`, `templates/manifest.json`,
  `tools/validate.py` (`validate_mls` rewritten for the 2026 files),
  `core/embedding.py` `SUITE_VERSION = "v1.13"`.

## Validation

`python tools/validate.py` must stay `RESULT: PASS`. The MLS run checks,
among others, `10,350 square feet` in the owner letter, `Block Application
No. 933` on the routing form and the written-out lot size in the ordinance.

The same change set was applied to the v1.10 code line as
`RCA_Wizard_Suite v1_10_2` (`scratchpad/port_mls_v10.py`: template paths via
`_TEMPLATES / name`, LibreOffice-only PDF conversion as in v1.10).
