# Suite v1.11 — developer notes

Companion to `MLS_WIZARD_DEVELOPER_MANUAL.md`. v1.11 changed how templates
are located, how PDFs are made, and how the suite is distributed. Wizard
generator code (the `generate_*` functions, context builders, token
replacement) is untouched; `tools/validate.py` passes the same 317 checks.

## New modules

| File | Purpose |
|---|---|
| `core/templates.py` | Template manifest: which `.docx` plays which role per wizard; replace / restore / status / baseline. Self-healing (`load_manifest` rebuilds from `DEFAULT_ROLES` when the JSON is missing or damaged). |
| `core/template_manager.py` | The Manage Templates screen (Tk frame hosted by `app.py`). UI only; all bookkeeping is in `core/templates.py`. |
| `core/office_convert.py` | `office_to_pdf()` and `recalc_with_excel()`. LibreOffice first, then Word/Excel via pywin32 COM. `RCA_OFFICE_ENGINE` forces one engine. |
| `core/launch.py` | Builds the command lines for helper processes (HCAD downloader, validate, onboard) in both the source layout and the frozen build. |
| `rca_tools.py` | Console entry point (`hcad` / `validate` / `onboard`). Becomes `RCA Wizard Tools.exe`. |
| `build/rca_wizard_suite.spec`, `build/build_package.py` | PyInstaller build; `Build Package.bat` drives it. |

## Exact edits to existing files

- **All four `wizards/*/wizard.py`**: the default template paths in the
  files dict now call `template_path(<wizard>, <role>, "<historical name>")`
  instead of `tmpl / "<historical name>"`; the HCAD subprocess call uses
  `core_launch.hcad_command()` / `hcad_cwd()` / `console_flags()` and the
  availability check uses `core_launch.hcad_available()`. Two new imports.
- **`wizards/mls/mailout_engine.py`, `wizards/mbl/mbl_mailout_engine.py`,
  both renewal wizards**: the body of `convert_office_to_pdf` /
  `_convert_office_to_pdf` delegates to `core.office_convert.office_to_pdf`.
  The old `_find_soffice` helpers remain (unused by the converter).
- **`wizards/mbl_renewal/wizard.py`**: the packet precheck asks
  `office_convert.available()` instead of `_find_soffice() is None`.
- **`core/recalc.py`**: when LibreOffice is absent, `recalculated_copy` tries
  `recalc_with_excel` before giving up.
- **`core/paths.py`**: when frozen, `WIZARD_FILES_DIR` is
  `<exe folder>/Wizard_Files` (the editable external folder).
- **`tools/validate.py`**: template names resolve through the manifest
  (`_t("<historical name>")`), new `--out DIR` option.
- **`tools/onboard_templates.py`**: reports go under `core.paths.WIZARD_FILES_DIR`.
- **`app.py`**: Manage Templates button + setup line on the home screen,
  `open_template_manager()`, `--templates` flag.
- **`core/embedding.py`**: `SUITE_VERSION = "v1.11"`.
- `requirements.txt`: `pywin32` (Windows only). `install_requirements.bat`:
  `--user`, offline `wheels` support. `tools/__init__.py` added so the tools
  import inside the frozen build.

## Frozen-build layout

```
RCA_Wizard_Suite_v1.11/
  RCA Wizard Suite.exe        windowed app (app.py)
  RCA Wizard Tools.exe        console helpers (rca_tools.py)
  _internal/                  Python runtime + libraries (never edit)
  Wizard_Files/
    templates/                editable; same layout as source, incl. manifest.json
    sample_input/  docs/  README.md
  README_FIRST.txt, Create Desktop Shortcut.bat, the two guides
```

Anything that needs a path at run time must go through `core.paths` (data)
or `core.launch` (helpers). `sys.executable` is the exe when frozen, so
never use it to run a `.py` file.

Playwright browsers are not bundled: the HCAD downloader connects to the
planner's own Chrome over CDP (`--cdp-url`), and the legacy fallback launches
Chrome with `channel="chrome"`.

## Known limits / next steps

- Three of the four wizards fill templates by matching the template's own
  placeholder wording. The Template Manager can install and archive those
  templates and run the harness, but it cannot detect a *changed placeholder
  phrase* (it will simply remain in the output). The reliable fix is to
  convert those templates and engines to `{{tokens}}`, one document at a time
  with the harness green before and after. Until then, always eyeball the
  sample output for wording templates.
- `core/recalc.py::needs_recalc` does not locate the Step 2 sheet in the MLS
  sample workbook (its relationship lookup returns nothing), so the poisoned
  cache repair only triggers for workbooks whose sheet can be found. The
  validation harness exercises it with the renewal workbooks. Pre-existing.
- The package is not code-signed; SmartScreen shows a one-time warning.
- Institutional names still hard-coded as fallbacks in engine code (director,
  former PC chair) are unaffected by the manifest.
