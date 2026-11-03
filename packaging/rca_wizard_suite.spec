# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for the City Planning RCA Wizard Suite.

Produces ONE folder containing two programs that share the same libraries:

    RCA Wizard Suite.exe   the windowed suite (home screen + four wizards)
    RCA Wizard Tools.exe   console companion: HCAD downloader, validation
                           harness, template onboarding reports

Templates, sample input and docs are NOT bundled here; build_package.py copies
them next to the executables as an editable Wizard_Files folder (see
core/paths.py, which looks there when the suite is frozen).

Build with build_package.py (or "Build Package.bat"), not by hand.
"""

from pathlib import Path
import re

SPEC_DIR = Path(SPECPATH).resolve()          # noqa: F821  (PyInstaller global)
SUITE_DIR = SPEC_DIR.parent
WF = SUITE_DIR / "Wizard_Files"
VERSION = re.search(r'SUITE_VERSION\s*=\s*"v([^"]+)"', (WF / 'core' / 'embedding.py').read_text(encoding='utf-8')).group(1)
ICON = str(SPEC_DIR / "rca.ico") if (SPEC_DIR / "rca.ico").exists() else None

# The wizards are imported lazily by name (app.py WIZARDS registry), so
# PyInstaller's static analysis cannot see them.  List them explicitly.
HIDDEN = [
    "tools.test_mbl_pc",
    "wizards.mbl_renewal.pc_speech",
    "wizards.mbl.pc_speech",
    "wizards.mbl.pc_report",
    "wizards.mbl.pc_data",
    "tools.test_mls_renewal_pc",
    "wizards.mls_renewal.pc_speech",
    "wizards.mls_renewal.pc_report",
    "wizards.mls_renewal.pc_data",
    "wizards.mls.pc_landuse",
    "wizards.mls.pc_map",
    "wizards.mls.pc_ui",
    "wizards.mls.pc_report",
    "wizards.mls.pc_data",
    "wizards.mls.pc_speech", "wizards.mls.pc_speech_ui", "tools.test_pc_speech",
    "core.ilms_holds", "core.ilms_browser", "core.ilms_ui", "tools.test_ilms_holds",
    "core.city_templates", "tools.test_city_templates", "core.mbl_templates_2026", "tools.test_mbl_templates_2026",
    "core.peer_review", "core.peer_readers", "core.peer_review_ui", "tools.test_peer_review",
    "core.designation_lookup", "tools.test_designation_lookup",
    "core.hcad_identity", "tools.test_hcad_duplicates", "core.automatic_case", "tools.test_simple_flow",
    "core.lot_divisions", "core.lot_review_ui", "core.lot_map_ui", "core.council_ui", "core.council_district",
    "tools.test_lot_upgrade", "tools.test_lot_upgrade_ui",
    "app", "core.startup", "core.property_letters",
    "tools.test_property_notifications",
    "tools.test_ordinance_counts",
    "tools.test_routing_forms", "core.routing_form", "tools.test_mbl_pc_approval",
    "wizards", "wizards.mls", "wizards.mls.wizard", "wizards.mls.mailout_engine",
    "wizards.mbl", "wizards.mbl.wizard", "wizards.mbl.mbl_mailout_engine",
    "wizards.mls_renewal", "wizards.mls_renewal.wizard",
    "wizards.mbl_renewal", "wizards.mbl_renewal.wizard",
    "core", "core.casefolder", "core.docx_unlock", "core.embedding", "core.errorlog",
    "core.hcad_downloader", "core.launch", "core.office_convert", "core.paths",
    "core.pdf_condense", "core.preflight", "core.recalc", "core.scrollable",
    "core.session", "core.sheet_layout", "core.soffice", "core.template_manager",
    "core.templates", "core.workbooks", "core.pdf_setup", "core.packet_work",
    "core.plat_search", "core.plat_search_ui", "core.plat_browser", "tools.test_plat_finder", "core.tax_map", "core.tax_map_ui", "tools.test_tax_maps",
    "core.boundary_map", "core.boundary_map_render", "core.boundary_map_ui",
    "core.application_ocr", "core.application_import", "core.application_import_ui",
    "tools.test_application_import", "tools.check_application_import", "tools.check_application_import_ui",
    "reportlab", "reportlab.pdfbase.ttfonts", "shapely", "numpy",
    "tools", "tools.validate", "tools.onboard_templates", "tools.repair_workbook",
    "tools.test_boundary_map", "tools.check_actual_boundary_maps",
    # Microsoft Office automation fallback (pywin32)
    "win32com", "win32com.client", "pythoncom", "pywintypes", "win32api",
    # stdlib pieces used indirectly
    "tkinter", "tkinter.ttk", "tkinter.filedialog", "tkinter.messagebox",
]

EXCLUDES = ["pytest", "IPython", "matplotlib", "scipy", "pandas", "PyQt5", "PySide2"]

block_cipher = None

suite = Analysis(
    [str(WF / "start_wizard.py")],
    pathex=[str(WF)],
    binaries=[],
    datas=[],
    hiddenimports=HIDDEN,
    hookspath=[],
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
)
tools = Analysis(
    [str(WF / "rca_tools.py")],
    pathex=[str(WF)],
    binaries=[],
    datas=[],
    hiddenimports=HIDDEN,
    hookspath=[],
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
)

suite_pyz = PYZ(suite.pure, suite.zipped_data, cipher=block_cipher)
tools_pyz = PYZ(tools.pure, tools.zipped_data, cipher=block_cipher)

suite_exe = EXE(
    suite_pyz, suite.scripts, [],
    exclude_binaries=True,
    name="RCA Wizard Suite",
    debug=False, bootloader_ignore_signals=False, strip=False, upx=False,
    console=False,
    icon=ICON,
)
tools_exe = EXE(
    tools_pyz, tools.scripts, [],
    exclude_binaries=True,
    name="RCA Wizard Tools",
    debug=False, bootloader_ignore_signals=False, strip=False, upx=False,
    console=True,
    icon=ICON,
)

COLLECT(
    suite_exe, suite.binaries, suite.zipfiles, suite.datas,
    tools_exe, tools.binaries, tools.zipfiles, tools.datas,
    strip=False, upx=False,
    name=f"RCA_Wizard_Suite_v{VERSION}",
)
