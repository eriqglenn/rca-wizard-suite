# Suite v1.12 — developer notes

Companion to `DEVELOPER_NOTES_v1_11.md`. v1.12 replaces the MLS Renewal
wizard's City templates with the 2026 set and rewrites
`wizards/mls_renewal/wizard.py` around them. The other three wizards are
untouched.

## What changed

- **Templates** (`templates/mls_renewal/`): SFR / No-SFR owner letters,
  applicant sign letter, Director approval, routing form and draft ordinance
  replaced by the 2026 versions (old files in `_previous/`, manifest history
  recorded). Three new roles: `protest_template` (protest / public-hearing
  letter to all owners), `staff_report_director_template` and
  `staff_report_pc_template` (OneMeeting NOVUS cover letters). The PC approval
  form is unchanged. `docs/mls_renewal/TEMPLATES_2026.md` has the field map.
- **Workbook** (`sample_input/mls_renewal/`): the City's 2026
  `RENEW_Min_Lot_Size_Calculation_Sheet_NEW.xlsx` with a new **Case Info** tab
  (built with Excel COM so the Step 2 logo, sensitivity label and formulas
  survive). `_v2.xlsx` is the blank planner copy, `_v2_FILLED_SAMPLE.xlsx` the
  demo case. The v1 workbooks were removed. Step 2 changed shape: the lot size
  to renew is `D2`, column F flags lots that meet it; there is no corner-lot /
  prevailing-lot-size block any more, so `_get_step2` and
  `_get_renewal_context` were rewritten (`Lots_Meeting_Minimum` replaces the
  old 75 %/10 % counts).
- **Merge engine**: placeholder matching is non-breaking-space aware (the
  2026 letters are full of `\xa0`); content controls are filled in place by
  their placeholder text and unwrapped (`_fill_content_controls`), checkbox
  controls are set by position (`_set_paragraph_checkboxes`); Next Steps
  numbering restarts per owner through a fresh `w:num` per letter
  (`_restart_list_numbering`) instead of the old text-prefix hack; the
  placeholder `mailto:` hyperlink is retargeted to the planner's e-mail;
  header/footer parts are reached through relationships, never through
  `section.header` / `section.footer` (python-docx would add a definition to
  every merged letter).
- **Packet**: 2026 order — routing form, approval, boundary map, tax map,
  original ordinance; plat, Step 2 staff analysis (checkbox) and HCAD records
  are optional extras. Staff-analysis export uses columns A:F of the 2026
  sheet.
- **UI**: template rows and buttons for the protest letters and the
  OneMeeting staff report (template chosen by `Approval_Path`); Review Owners
  lists the owners; Case Info warns when `Protest_Count` > 0 but the approval
  path is still Director.
- `core/templates.py` `DEFAULT_ROLES` (3 new roles), `templates/manifest.json`,
  `core/preflight.py` snapshot fields, `tools/validate.py` (v2 workbook,
  protest/staff-report checks, section-shape check for multi-section letter
  templates), `core/embedding.py` `SUITE_VERSION = "v1.12"`.

## Validation

`python tools/validate.py` must stay `RESULT: PASS`. The renewal run
generates every 2026 document from the filled sample and checks, among
others, `Previous Ordinance: 2006-118` on the routing form, `5,950 square
feet` on the Director approval and the written-out lot size in the ordinance.

The same change set was applied to the v1.10 code line as
`RCA_Wizard_Suite v1_10_1` (no manifest there: `template_path` is a shim that
resolves the historical file names; PDF conversion still needs LibreOffice as
in v1.10).
