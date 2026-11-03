# RCA Wizard Suite v1.0 — QA Report

Date: 2026-06-10
Scope: merge of the four standalone wizards into one suite. Generator code is
unchanged from the standalone builds; only the window embedding, template /
sample path resolution, and the shared HCAD downloader location changed.

## Source builds merged

| Wizard | Standalone build | Code taken from |
|---|---|---|
| MLS | v14.4 (package v14.5) | mailout_engine.py + rca_mailout_wizard.py |
| MBL | v8.0 (package v8.1) | mbl_mailout_engine.py + mbl_wizard.py |
| MLS Renewal | v3.0 | mls_renewal_wizard.py |
| MBL Renewal | v1.0 (package v1.3) | mbl_renewal_wizard.py |

`hcad_downloader.py` was byte-identical (same MD5) across all four packages;
the suite keeps one copy in `core/`.

## Exact changes made to wizard code (per wizard)

1. Class declaration: `tk.Tk` -> `EmbeddedWizardMixin, tk.Frame`
   (`__init__(self, parent=None)`; the mixin forwards title/geometry/minsize
   to the shared toplevel — no other window-manager calls exist in any wizard).
2. Template defaults: resolve via `core.paths.templates_dir(<key>)` instead of
   `Path.cwd()` (MLS) / module-local `templates` folder (others).
3. Sample-input defaults (renewals): resolve via `core.paths.sample_input_dir`.
4. HCAD launch path: the single shared `core/hcad_downloader.py`.
5. `if __name__ == "__main__"` blocks replaced with launch guidance
   (wizards start through `app.py` / `Run Wizard.bat`).
6. MLS engine import made package-relative; MLS CLI given a path bootstrap.

Nothing else in any engine or wizard module was modified.

## Validation performed (per master brief §20.12)

### Structural (XML-level), on documents generated from the bundled samples

- Notification letters, preview + all-owners, all four wizards (SFR and
  No-SFR for MLS Renewal): **zero** leftover `{{...}}` tokens anywhere.
- §20.6a section/footer balance on combined mail-outs:
  - MLS all-owners: 56 sections, titlePg=28, footerReference=28 (28 owners) —
    the exact balanced signature required by the brief.
  - MBL, MLS Renewal (SFR + No-SFR), MBL Renewal combined outputs: section
    counts match owners; footer references present on every letter section.
- Routing forms, all four wizards: checkbox state and glyph verified
  (MLS/MBL/MBL Renewal: w14:checked values + ☒/☐ glyphs both set;
  MLS Renewal writes glyphs after stripping content controls — its original
  tested design). Zero leftover tokens.
- Director approvals, PC approvals, applicant letters, draft ordinances:
  generated for all four wizards, zero leftover tokens.

### Live UI smoke test (virtual display)

- Suite home screen renders; each wizard opens embedded, navigates all
  four steps (Select Files / Case Info / Owners / Generate), and returns
  Home with full state reset; re-opening works.
- Title bar shows both builds, e.g.
  `City Planning RCA Wizard Suite v1.0 — MLS Renewal Planner Wizard v3.0`.

### Packaging

- `python -m compileall`: clean. All modules import headlessly.
- `Run Wizard.bat` purges every `__pycache__` under Wizard_Files before launch.
- No sample_output folders shipped (per §20.9).

## Known pre-existing notes (not introduced by the merge, left untouched)

- The renewal monoliths emit an lxml `FutureWarning` ("Truth-testing of
  elements") at one helper; harmless today, worth fixing during the planned
  shared-core extraction.
- The monoliths contain later redefinitions of some generator functions
  (`generate_routing_form` etc. appear more than once; Python uses the last
  definition). Identical to the standalone builds.

## Outstanding (recommended next phase)

- Shared-core extraction: unify the two renewal monoliths (~80% identical)
  onto `core/` modules, then fold MBL-new onto the MLS engine structure.
- Final environment check on a planner Windows machine with Microsoft Word
  (§20.12.3) before department-wide rollout — required for every release.


---

# v1.3 additions (validated)

- **Scroll fix (v1.1):** shared ScrollableContent wrapper; removed the three
  nested page canvases that clipped lower buttons. Verified at 600px window:
  all four generate pages scroll, RCA Packet button reachable.
- **MBL Renewal recolored orange (v1.1)** including the home-screen card.
- **Map gate for notification letters (v1.2):** both renewal wizards now
  block notification letters without a boundary map (matching MLS/MBL);
  verified that director/PC/routing/applicant still generate with no map and
  no attachments.
- **tools/validate.py:** executable §20.12 validation contract — 70 checks
  across all four wizards (placeholders, section/footer preservation,
  checkboxes, file outputs). Full run: PASS. The balance check is
  structure-aware per §20.6a: titlePg==footerRef for first-page-footer
  templates (MLS), footerRef>=sections for ordinary-footer templates.
- **tools/onboard_templates.py:** §20.0 onboarding reports generated for all
  27 templates into docs/onboarding (diff baseline for City revisions).
- **Centralized error logging:** all generator failures append a full
  traceback to wizard_error_log.txt in the output folder; the error dialog
  names the log file. Verified by forcing a failure.
- **Preflight panel** on every Case Info page (non-blocking ✓/⚠ checklist);
  renewal Case Info pages additionally show a Case snapshot grid (previously
  those pages displayed no data — brief §5). Verified rendering in all four.
- **Per-case output folders:** Documents/RCA_Cases/<app number> suggested on
  workbook load; junk app numbers (0, prose fragments) rejected with
  workbook-name fallback; custom folders never overridden. Verified.
- **Session memory:** last file selections restored per wizard across
  Home/reopen; stale paths skipped. Stored in ~/.rca_wizard_suite_session.json.
  Verified across reopen.


---

# v1.5 additions (validated)

- **Step 1 -> Step 2 data transfer fixed at the source:** the bundled MLS
  Renewal FILLED sample shipped with a poisoned formula cache (every formula
  saved with an empty cached value and no recalc flag), so Step 2 displayed
  blank in Excel and read as empty by the wizard, despite correct formulas and
  data. Replaced with a recalculated copy carrying real cached values
  (verified: SFR 86%, renewal 80%, 21 lots, 16 qualifying, min 5,950 sq ft
  now flow into generated approvals/letters). The v1 blank starter was
  cache-repaired with fullCalcOnLoad.
- **Wizard self-heal for poisoned workbooks:** all four wizards now resolve
  the workbook through a lazy `_wb()` that detects an empty Step 2 cache
  (core/recalc.needs_recalc) and reads from a recalculated temporary copy
  (LibreOffice, isolated profile with plain fallback). The planner's file is
  never modified; case-folder naming and preflight intentionally keep using
  the original path.
- **Letter spacing template fidelity:** fixed the renewal wizards'
  `_set_run_text_preserve_xml` to be a true inverse of run.text — existing
  <w:tab/>/<w:br/> elements are removed and tab/newline characters are written
  back as proper elements. This removes the doubled tab that pushed the
  notification date right of the template position (user-reported on MBL
  Renewal). MBL-new address indent aligned to the template's 0.5" (was 0.55").
  The No-SFR subject continuation alignment remains the documented deliberate
  fix.
- **Spacing checks added to the validation contract:** tools/validate.py now
  asserts, for the head region of every letter type, that each paragraph's
  effective first-line position and tab structure match the template (with the
  documented No-SFR allowance). Full run: 186/186 checks PASS.


---

# v1.6 additions (validated)

- **Minimum value missing from letters (user-reported):** on a workbook with
  an empty formula cache the renewal letters printed "the minimum lot size is
  square feet." because (a) the value was computed from Step 2 mirror-formula
  cells that read blank, and (b) `.get(key, default)` does not default on an
  EMPTY string, so the placeholder was replaced with nothing. Fixed both
  layers in all four wizards:
  - Renewals derive the most-frequent lot size / building line from the raw
    Step 1 columns their Step 2 mirrors read (identical multisets; no City
    calculation re-implemented) when the cache is empty. Verified: the
    user's original poisoned workbook now yields "5,950 square feet" in SFR
    and No-SFR letters; poisoned MBL Renewal yields "35 feet".
  - New-application minimums are genuine 70%-rule calculations (Step 2
    K16/L16) that the wizard must never re-implement, so an empty value now
    keeps the template's visible X,XXX / X placeholder instead of silently
    blanking; the automatic workbook recalc remains the value-recovery path.
- **Validation contract extended (216 checks):** every generated letter is
  scanned for blank-merge artifacts ("is  square feet", "of  feet", etc.),
  and both renewal validators regenerate the notification from a deliberately
  cache-poisoned copy of the sample and assert the minimum value still lands.


---

# v1.7 additions (validated)

- **Date timestamp leak fixed (user-reported):** Excel stores real dates as
  datetime objects; the renewal wizards' _clean() and the MBL-new clean()
  stringified them with str(), leaking a midnight timestamp into documents
  (e.g. the ordinance WHEREAS clauses: "effective on 2006-01-18 00:00:00").
  All four wizards now format datetimes as the City long form
  ("January 18, 2006"), built from parts (no %-d/%#d) so Windows and Linux
  agree. Plain-string dates, numbers, and blanks pass through unchanged. MLS
  new already handled dates; its formatter was hardened to the same parts
  approach. Verified across MLS/MBL renewal ordinance, director and PC
  approvals, and notification letters using a workbook whose Case Info dates
  are real datetimes: no timestamps, correct "Month D, YYYY" output.
- **Validation contract extended (274 checks):** every renewal validator now
  rewrites Case Info date fields as real datetimes and asserts the generated
  ordinance, director approval, and notification contain no leaked timestamp
  or bare ISO date (body, headers, and footers). MBL new gets the same check
  on its letter.


---

# v1.8 additions (validated)

- **Generated documents not editable (user-reported):** several City templates
  ship with Word document protection turned on
  (`<w:documentProtection w:edit="forms" w:enforcement="1">`, with a password
  hash) -- the MLS and MBL ordinance templates and all four routing-form
  templates. Word then locks the whole document so only form fields can be
  typed into; the surrounding text is read-only. The wizards copied the
  template faithfully, so every generated ordinance/routing form inherited the
  lock.
- **Fix:** a shared core/docx_unlock.py removes document/write protection from
  the OUTPUT file (never the template on disk), and every public generator in
  all four wizards calls it as the last step before returning. Per the
  decision to make every Word doc fully editable, this runs on all generated
  documents, not just ordinances. Routing-form checkboxes are unaffected
  because the wizard writes both the w14:checked state and the visible ☒/☐
  glyph into the OOXML before the unlock; verified glyph counts unchanged
  after unlocking.
- **Validation contract extended (304 checks):** check_docx now asserts every
  generated document is free of enforced document/write protection, so a
  future template that arrives protected (or a generator that forgets to
  unlock) fails the build.


---

# v1.9 additions (validated)

- **Planning Commission chair updated (user-reported):** the new PC chair is
  Lisa Clark; Sonny Garza remains Vice-Chair. Per the instruction to treat the
  chair as template content (not wizard-injected case data), the change was
  made in the document files only:
  - All four PC approval templates now display "Lisa Clark, Chair" (was
    "Martha L. Stein, Chair" / "Martha Stein, Chair" in three; MLS Renewal
    showed "Lisa M. Clark, Chair" and was aligned to drop the middle initial).
    "Sonny Garza, Vice-Chair" is unchanged. Text-node-only edits; all run
    formatting preserved; templates remain valid.
  - The PC_Signer value in the sample workbooks' Case Info tabs was updated to
    "Lisa Clark, Chair" so a planner reusing a sample does not reintroduce the
    former chair. These edits were made surgically in the cell XML so the
    formula cache is NOT disturbed (no return of the section 20.4 stale-cache
    problem); verified each workbook differs from the prior build by exactly
    one cell (the PC_Signer value) with all formulas byte-identical.
  - Wizard name-handling code was intentionally left unchanged.
- **Validation contract extended (316 checks):** every generated PC approval is
  asserted to name the current chair (Lisa Clark, Chair) and vice-chair
  (Sonny Garza, Vice-Chair) and to contain no former-chair name, so a future
  template revision that reintroduces the old name fails the build.


---

# v1.10 additions (validated)

- **Qualifying-lot count corrected (user-reported):** the Director / PC /
  ordinance sentence "N lots ... are at least X square feet in size" was
  pulling N from the 70%-cumulative-area rank (19 in the sample), which is the
  minimum-lot-size *determination* rank, not the count of lots that meet the
  minimum. Per the decision to treat this as planner-owned data, the wizard no
  longer calculates it: it reads a new planner-entered Case Info field
  Lots_Meeting_Minimum (the planner reads the count off the Step 2 "lots ranked
  by size" table). All MLS-new documents that use the Qualifying_Lot_Count
  tokens (Director approval, PC approval, draft ordinance) now show the
  planner's value (28 in the sample, verified). If the field is blank, the
  documents show a visible "X" placeholder instead of a wrong number or a raw
  {{token}}. The field was added to the sample Case Info tab surgically (cell
  XML), so the formula cache is not disturbed.
  - The former 70%-rank computation was removed as a source of
    Qualifying_Lot_Count in every active builder; the genuine 70%-area
    percentage (the "representing X%" phrase) is retained.
  - Qualifying_Lot_Count was removed from the required-field validators (it is
    now optional/planner-entered), and from the routing-form validator where it
    was never used by the template.
  - MLS Renewal is unaffected: its approvals use the different prevailing-lot-
    size criterion (75% exclusive of corner lots), not this count.
- **Validation contract extended (317 checks):** asserts the Director approval
  qualifying sentence carries the Case Info value (28), not the 70%-rank (19).
