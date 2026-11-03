# City Planning RCA Wizard Suite v1.26

One program containing all four City Planning RCA document wizards:

| Wizard | Build | Purpose |
|---|---|---|
| MLS Wizard | v15.0 | Special Minimum Lot Size — new block applications |
| MBL Wizard | v9.0 | Special Minimum Building Line — block applications |
| MLS Renewal Wizard | v3.0 | Special Minimum Lot Size — block renewals |
| MBL Renewal Wizard | v2.0 | Special Minimum Building Line — block renewals |


## Selecting case files

**Create Boundary Map** is available in Required setup in all four wizards.
Save the calculation sheet in Excel, select that workbook and an output folder,
then click the button. The map uses the sheet's property addresses/accounts,
case description, minimum value and case number.
The selected wizard determines MLS, MBL, MLS Renewal or MBL Renewal wording.
All generated maps include the supplied City of Houston seal and Planning &
Development Department wordmark in the lower-right legend, below the disclaimer.
You can recreate older maps with **Create Boundary Map** to add the seal.
Older or mismatched maps show a warning only; you can generate documents with
the map you selected without updating it.
Every listed property stays inside the map boundary, including properties
classified EXC. Land-use labels and the land-use key are omitted. Each address
stays inside its own matched parcel; unaddressed lots remain blank.
Block ranges automatically use a hyphen, for example `100-200 blocks of Main
Street`, while wording such as `north and south sides` remains unchanged.

Missing City street names now trigger a Harris County street lookup. When a
street still cannot be located, the preview shows a warning and a movable
street label; **Use This Map** remains available. Step 2 supplies the map
description, and the saved provenance records warnings and alternate sources.

Review the preview. Orientation, framing and label positions can be adjusted;
street labels keep their standard font size. Small-lot addresses rotate or use
a smaller font to fit inside the parcel. **Use This Map** saves a vector PDF,
a 200-dpi PNG and a small provenance file, and selects the PDF for letters and
the RCA packet. Previous map files are retained. A changed workbook or wizard
type requires regenerating an attached generated map before it can be used in
documents. Existing externally prepared maps can still be selected with Browse.

Internet access is required to retrieve public HCAD parcels and City street
data. Data is cached in `.boundary_map_cache` inside the selected output folder;
the preview shows retrieval dates. **Refresh parcel data** fetches fresh data.
No desktop GIS, browser automation or GIS license is required for this feature.
Windows Arial and Office Calibri fonts must be available for the City layout.

Stale accounts can be resolved by an exact address and legal-description match;
the preview reports the replacement account. Unaddressed parcels use their
account and legal description. Houston Map Viewer provides retained parcel
outlines when an old account is missing from the county layer. Unresolved
ambiguity, unknown land uses and incomplete service results still stop the map.
The default MBL instructional sheet creates an explicitly labeled Haver Street
example. Use your saved calculation sheet for an application map.
See `docs/BOUNDARY_MAPS.md` for details and known limitations.

Type or paste a path into a file field, or use **Browse**. The displayed path is
now the path used by generation in all four wizards. Clearing an attachment
removes the previous selection. This also applies to the calculation workbook,
output folder and one-off template overrides. Save workbook edits in Excel
before refreshing. Downloaded calculation sheets keep their examples and formulas.

## RCA packet temporary folders

Each packet attempt uses a new folder in the Windows temporary directory.
An old `_renewal_packet_work`, `_mbl_packet_work` or `_packet_work` folder in
the case directory is left alone; it does not need to be deleted before retrying.
The final RCA still saves to the output location you selected. If Windows locks
the new temporary folder during cleanup, the completed packet still succeeds.

## PDF conversion on work computers

Use **Check PDF setup** on the home screen to create temporary test PDFs with
**desktop Word and Excel**. The report identifies the launch mode and preserves
any actual startup/export error. Use **Copy report** or **Save report** when
asking for help. The check does not open case files.

A failed registry lookup no longer stops the wizard from trying Word or Excel.
Document PDFs use Word; spreadsheet PDFs use Excel. LibreOffice remains an
alternative. Browser-only Office cannot perform these desktop exports. The
Windows package includes the Office automation components; a source install
requires the Windows dependencies listed in `requirements.txt`.

## Calculation sheets and current templates

Each home-screen card has **Download calculation sheet** and **Open wizard**.
Download saves an editable `.xlsx` copy where you choose and opens it in Excel.
The examples, instructions, colors, dropdowns and formulas are preserved. Replace
example entries with your application's information as you work. The bundled
file is never edited by downloading.

Each wizard starts with its current instructional calculation sheet unless you
previously selected a separate case workbook. Browse to your saved case copy
when you are ready to load it. Document templates always default to the active
installed set for that application type. One-off Browse overrides apply only
while that wizard is open; they no longer replace the defaults next time.
The existing MBL ordinance template remains in use because the 2026 replacement
is encrypted. This release does not change City templates or calculation rules.

After editing the calculation sheet, **save it in Excel**, then click **Load
Workbook** (or **Refresh from Workbook** in MLS). Review and generation compare
the saved file's contents and discard an older recalculated copy after any
change. Unsaved Excel edits cannot be read by the wizard.


## How to start

**Packaged build (planners, City computers):** extract
`RCA_Wizard_Suite_v1.26_Windows.zip` to a folder you own and double-click
**`RCA Wizard Suite.exe`**. No Python, no setup, no admin rights. Step by
step: `docs/INSTALL_CITY_COMPUTERS.md`.

**From source (developer computer with Python):**

1. **First time only:** double-click `install_requirements.bat`.
2. Double-click **`Run Wizard.bat`**.

Then, either way:

3. On the home screen, click the wizard for the application you are processing.
4. Use the **← All Wizards** bar at the top to return to the home screen.
   Returning home fully resets the wizard, so each case starts clean.

> Extract every new suite package into a **fresh, empty folder** — never on top
> of an old one. `Run Wizard.bat` also clears stale compiled code on every
> launch.

## Where things live

```
RCA_Wizard_Suite/
  Run Wizard.bat              <- start here (source install)
  install_requirements.bat    <- first-run setup (source install; uses an
                                 offline "wheels" folder when one is present)
  Build Package.bat           <- makes the no-Python zip for other computers
  build/                      PyInstaller spec + build script
  tools/make_offline_wheels.bat
  Wizard_Files/
    app.py                    suite home screen / launcher
    rca_tools.py              console helpers: validate | onboard | hcad
    core/                     shared code (HCAD downloader, PDF engine,
                              template manifest, Template Manager, ...)
    wizards/                  the four wizards (generator code unchanged from
                              the standalone builds, embedded in the suite)
    templates/
      manifest.json           which .docx plays which role for each wizard
                              (edited for you by Manage Templates)
      mls/  mbl/  mls_renewal/  mbl_renewal/
                              official City templates, one folder per wizard.
                              When the City revises a template, use Manage
                              Templates on the home screen (or drop the file
                              here and edit manifest.json). Replaced files are
                              kept in <wizard>/_previous. No code change
                              needed. Every wizard also has Browse buttons for
                              one-off overrides. See docs/TEMPLATE_UPDATE_GUIDE.md.
    sample_input/             sample workbooks and maps, per wizard
    docs/                     per-wizard build notes, field maps, and the
                              install / template guides
```

## What the suite does for you automatically

- **Per-case output folders** — when you load a workbook, the output folder
  rolls to `Documents\RCA_Cases\<application number>\` (or the workbook name
  if no number is set), so a case's letters, forms, HCAD records, and packet
  stay together. A folder you picked by hand is never overridden.
- **Session memory** — each wizard remembers your last file selections
  (case workbook, map, attachments, output folder) between sessions.
  Stale paths are skipped automatically.
- **Preflight checks** — the Case Info page shows a non-blocking checklist
  (workbook tabs, owner-row count, required Case Info fields, map/output
  selection). Nothing on this list ever blocks generation; each document
  still requires only the files it actually uses.
- **Error log** — if a generator fails, the full details are appended to
  `wizard_error_log.txt` in the output folder. Include that file when
  reporting a problem.

## Updating templates and checking them

Click **Manage Templates** on the home screen. Pick the wizard, select the
document, **Replace selected template...**, and let it run the check. The
previous file is archived and can be restored with one click. Details and the
manual procedure: `docs/TEMPLATE_UPDATE_GUIDE.md`.

## PDF engine

RCA packets and staff-analysis PDFs are produced by LibreOffice when it is
installed, otherwise by Microsoft Word/Excel (standard on City computers).
The same applies to repairing a workbook whose Step 2 cache is blank. The
home screen shows which engine is active. Force one with the environment
variable `RCA_OFFICE_ENGINE=libreoffice` or `RCA_OFFICE_ENGINE=msoffice`.

## Making the shareable package

Double-click **`Build Package.bat`** (needs Python 3.13 with `pyinstaller`,
`pywin32`, and the requirements installed once). The result,
`dist\RCA_Wizard_Suite_v1.26_Windows.zip`, runs on any Windows PC with
Microsoft Office and Chrome. Step by step: `docs/INSTALL_CITY_COMPUTERS.md`.

## Build tools (for the wizard maintainer)

From the `Wizard_Files` folder:

```
python tools\validate.py             # run the full §20.12 validation contract
python tools\validate.py --wizard mls --out C:\temp\proofs   # keep the documents
python tools\onboard_templates.py    # regenerate template onboarding reports
```

`validate.py` generates every document type from the bundled samples and
asserts the structural checks (no unfilled placeholders, section/footer
preservation in combined mail-outs, checkbox glyphs). Exit code 0 = PASS.
`onboard_templates.py` writes one §20.0 report per template into
`docs\onboarding`; when the City revises a template, re-run it and diff
against the saved report to see exactly what changed.

In the packaged build the same tools run through
`"RCA Wizard Tools.exe" validate|onboard|hcad`.

## Version diagnostics

The window title always shows **both** the suite build and the active wizard
build, e.g.:

```
City Planning RCA Wizard Suite v1.26 — MLS Renewal Planner Wizard v3.0
```

When reporting an issue, include the full title-bar text and the "PDF engine"
line from the home screen.

## Direct launch (optional)

```
python Wizard_Files\app.py --wizard mls
python Wizard_Files\app.py --wizard mbl
python Wizard_Files\app.py --wizard mls_renewal
python Wizard_Files\app.py --wizard mbl_renewal
python Wizard_Files\app.py --templates        # Template Manager
```

The same flags work on `RCA Wizard Suite.exe` in the packaged build.

## What changed in v1.19

- Fixed typed/pasted file fields being ignored by MBL, MLS Renewal and MBL Renewal.
  Previously only Browse updated the paths used by generation; the screen could
  display a different workbook or output folder from the one actually used.
- Fixed eight MLS delayed error callbacks losing the original exception before
  the error dialog/log could receive it.
- An empty output folder now reports a selection error instead of saving into
  the application directory.
- Error logs now include suite version, EXE/source launch mode and runtime path.
- Added seven regressions which were run against the old behavior first. They cover
  actual editable widgets, navigation, workbook switching, clearing attachments,
  Browse/typing combinations, packet generation and delayed error reporting.
- Templates, calculations, packet contents and Office conversion remain unchanged.
  See `docs/DEVELOPER_NOTES_v1_19.md` for the reproduction and
  the distinction between development tests and the work-computer status.

## What changed in v1.18

- Fixed RCA generation stopping with Access denied while deleting an earlier
  packet's working folder, before Word/Excel conversion could begin.
- All four wizards now use a separate system temporary folder for each packet.
  They leave previous working folders in the case directory untouched.
- Cleanup errors cannot hide a conversion error or fail a completed packet.
- Added five packet-workspace regression tests, including repeated generation
  in all four wizards with blocked old folders. Templates and packet order remain
  unchanged. See `docs/DEVELOPER_NOTES_v1_18.md`.

## What changed in v1.17

- Corrected Office detection: try actual Word/Excel automation even when Windows
  installation hints are missing. Excel exports no longer require Word detection.
- Preserve conversion errors instead of reporting every failure as missing Office.
- Added **Check PDF setup** with separate, bounded Word and Excel export tests and
  a copyable/savable report. The packaged test uses the companion EXE, not Python
  installed on the work computer.
- Applied the shared Office fallback to the spare source line as well.
- Preserved calculation sheets, City templates, document content and packet order.

## What changed in v1.16

- Added a calculation-sheet download button to each home-screen application card.
- Preserved the instructional workbooks byte for byte, including their examples.
- Defaulted all four wizards to the current instructional workbook and active
  installed document templates; older remembered template selections are skipped.
- Fixed stale recalculated workbook reuse after a save at the same source path.
  All four wizards share the content-aware read cache.
- Added `python tools/test_workflow.py` for the download/default/refresh regressions.

## What changed in v1.15

- **Address blocks print without the grey field shading** the City templates
  leave on the «merge field» runs; applies to every letter in all four wizards.
- **Owner letters: the boundary map is scaled to fit the page body** of the map
  section (a tall map used to overflow and add a blank page after each
  letter), and the section break between mail-out letters no longer takes a
  line of its own. Applies to all four wizards.
- **MBL Wizard v9.0 uses the 2026 City templates** (owner letter, applicant
  sign letter, Director approval, Planning Commission approval, routing form)
  and six new documents: the notice of public hearing (support under 51%), the
  protest letter to all owners, the notice of Planning Commission rescheduling
  and three OneMeeting staff reports (Director path, under 51%, protest). The
  wizard picks the staff report from Case Info. The previous MBL templates are
  archived in `templates\mbl\_previous`.
- **New MBL calculation workbook** with a Case Info tab
  (`sample_input\mbl\MBL_Calculation_Sheet_Case_Info_v3.xlsx`, plus a filled
  sample). The City's 2026 Step 1 sheet carries an extra HCAD_NUM_1 column.
- The MBL draft ordinance still uses the previous token template: the City's
  2026 ordinance file is password-protected and cannot be opened.
- Field map and placeholder wording: `docs\mbl\TEMPLATES_2026.md`. Developer notes: `docs\DEVELOPER_NOTES_v1_15.md`.

## What changed in v1.14

- **MBL Renewal Wizard v2.0 uses the 2026 City templates** (owner letter,
  applicant sign letter, Director approval, routing form, draft ordinance) and
  four new documents: the protest letter to all owners, the notice of Planning
  Commission rescheduling, the application acceptance letter and the OneMeeting
  staff report. The previous MBL Renewal templates are archived in
  `templates\mbl_renewal\_previous`.
- **New MBL Renewal calculation workbook** with a Case Info tab
  (`sample_input\mbl_renewal\MBL_Renewal_Calculation_Sheet_Case_Info_v2.xlsx`,
  plus a filled sample). Pre-2026 case workbooks still load.
- Field map and placeholder wording: `docs\mbl_renewal\TEMPLATES_2026.md`. Developer notes: `docs\DEVELOPER_NOTES_v1_14.md`.

## What changed in v1.13

- **MLS Wizard v15.0 uses the 2026 City templates** (owner notification letter
  with HCAD legal description and land-use confirmation, applicant sign letter,
  Director approval, routing form, draft ordinance) and three new documents:
  the protest letter to all owners, the notice of public hearing for support
  under 51 %, and the OneMeeting staff report. The previous MLS templates are
  archived in `templates\mls\_previous`.
- **New MLS calculation workbook** with a Case Info tab
  (`sample_input\mls\MLS_Calculation_Sheet_Case_Info_v13.xlsx`, plus a filled
  sample). Step 2 layout is unchanged, so v12 case workbooks still load.
- **MLS RCA packet** follows the 2026 routing form: routing form, approval,
  boundary map, tax map; plat, Step 2 staff analysis and HCAD records are
  optional extras.
- Field map and placeholder wording: `docs\mls\TEMPLATES_2026.md`. Developer notes: `docs\DEVELOPER_NOTES_v1_13.md`. `tools\validate.py` passes all checks.

## What changed in v1.12

- **MLS Renewal Wizard v3.0 uses the 2026 City templates** (owner letters,
  applicant sign letter, Director approval, routing form, draft ordinance) and
  three new documents: the protest / public-hearing letter to all owners and
  the OneMeeting staff reports (Director path and Planning Commission path).
  The previous renewal templates are archived in
  `templates\mls_renewal\_previous`.
- **New renewal calculation workbook** with a Case Info tab
  (`sample_input\mls_renewal\MLS_Renewal_Calculation_Sheet_Case_Info_v2.xlsx`,
  plus a filled sample). Yellow = planner enters, green = dropdown, blue = read
  from Step 2. Copy the blank workbook for each new case.
- **Renewal RCA packet** follows the 2026 routing form: routing form,
  approval, boundary map, tax map, original ordinance; subdivision plat, Step 2
  staff analysis and HCAD records are optional extras.
- Field map and placeholder wording: `docs\mls_renewal\TEMPLATES_2026.md`.
  Developer notes: `docs\DEVELOPER_NOTES_v1_12.md`. `tools\validate.py`
  passes all 360 checks.

## What changed in v1.11

- **Manage Templates** screen and `templates/manifest.json`: replace any
  City template without renaming it or editing code; old files archived in
  `_previous`; one-click restore; built-in check and sample-document
  generation from the screen.
- **No LibreOffice required**: PDF conversion and workbook recalculation fall
  back to Microsoft Word/Excel (`core/office_convert.py`).
- **Packaged build** (`Build Package.bat`, `build/`): a zip that runs with no
  Python installed. Helpers run through `RCA Wizard Tools.exe`
  (`rca_tools.py`); `core/launch.py` picks the right command in both layouts.
- `install_requirements.bat` installs per-user and from an offline `wheels`
  folder when present (`tools/make_offline_wheels.bat`).
- `tools/validate.py --out DIR` keeps the proof documents in a chosen folder.
- Wizard generator code is otherwise unchanged; `tools/validate.py` still
  passes all 317 checks.

## What changed vs. the four standalone packages

- One launcher and home screen; pick the wizard per case.
- One shared `hcad_downloader.py` in `core/` (the four copies were identical;
  fixes now land once for all wizards).
- Templates and sample input are centralized under `Wizard_Files\templates`
  and `Wizard_Files\sample_input`, namespaced per wizard.
- Each wizard's generators, templates, field mappings, and outputs are
  otherwise **unchanged** from the standalone builds listed above.
