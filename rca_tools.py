"""Console companion for the RCA Wizard Suite.

Runs the helper programs that the wizards start as separate processes.  From
source this is optional (the wizards can call the scripts directly); in the
packaged build it becomes ``RCA Wizard Tools.exe`` and is the only way to run
them, because the package contains no Python interpreter.

    rca_tools.py hcad     [downloader arguments]   HCAD record downloader
    rca_tools.py validate [--wizard KEY] [--out DIR]  validation harness
    rca_tools.py onboard  [--wizard KEY]            template onboarding reports
    rca_tools.py pdfcheck --app word|excel --out DIR  desktop Office PDF test
    rca_tools.py mapcheck                           boundary-map self-test
    rca_tools.py mapcasecheck --cases FILE --out DIR actual workbook/UI checks
"""

from __future__ import annotations

import sys
from pathlib import Path

_WIZARD_FILES = Path(__file__).resolve().parent
if str(_WIZARD_FILES) not in sys.path:
    sys.path.insert(0, str(_WIZARD_FILES))

USAGE = __doc__


def main() -> int:
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help", "help"):
        print(USAGE)
        return 2
    cmd = sys.argv.pop(1).lower()
    if cmd == 'mblpc_check':
        import unittest
        from tools import test_mbl_pc
        tests = unittest.defaultTestLoader.loadTestsFromModule(test_mbl_pc)
        return 0 if unittest.TextTestRunner(verbosity=2).run(tests).wasSuccessful() else 1
    if cmd == 'renewalpc_check':
        import unittest
        from tools import test_mls_renewal_pc
        tests = unittest.defaultTestLoader.loadTestsFromModule(test_mls_renewal_pc)
        return 0 if unittest.TextTestRunner(verbosity=2).run(tests).wasSuccessful() else 1
    if cmd == 'citytemplatecheck':
        import unittest
        from tools import test_city_templates
        tests = unittest.defaultTestLoader.loadTestsFromModule(test_city_templates)
        return 0 if unittest.TextTestRunner(verbosity=2).run(tests).wasSuccessful() else 1
    if cmd == 'peercheck':
        from tools.test_peer_review import main as check
        return check()
    if cmd == 'designationcheck':
        from tools.test_designation_lookup import main as check
        return check()
    if cmd == 'simpleflowcheck':
        from tools.test_simple_flow import main as check
        return check()
    if cmd == 'hcaddedupcheck':
        from tools.test_hcad_duplicates import main as check
        return check()
    if cmd == 'lotcheck':
        from tools.test_lot_upgrade import main as check
        return check()
    if cmd == 'lotuicheck':
        from tools.test_lot_upgrade_ui import main as check
        return check()
    if cmd == 'speechcheck':
        from tools.test_pc_speech import main as check
        return check()
    if cmd == 'pcreportcheck':
        from tools.test_pc_report import main as check_pc_report
        return check_pc_report()
    if cmd == 'pcapprovalcheck':
        from tools.test_mbl_pc_approval import main as check_pc_approval
        return check_pc_approval()
    if cmd == 'condensecheck':
        from tools.test_pdf_condense import main as check_condense
        return check_condense()
    if cmd == 'platcheck':
        from tools.test_plat_finder import main as check_plats
        return check_plats()
    if cmd == 'taxmapcheck':
        from tools.test_tax_maps import main as check_tax_maps
        return check_tax_maps()
    if cmd == 'mbltemplatecheck':
        from tools.test_mbl_templates_2026 import main
        return main()
    if cmd == 'ilmscheck':
        from tools.test_ilms_holds import main
        return main()
    if cmd == 'routingcheck':
        from tools.test_routing_forms import main as check_routing_forms
        return check_routing_forms()
    if cmd == 'ordinancecheck':
        from tools.test_ordinance_counts import main as check_ordinances
        return check_ordinances()
    if cmd == 'lettercheck':
        from tools.test_property_notifications import main as check_property_letters
        return check_property_letters()
    if cmd == 'importuicheck':
        from tools.check_application_import_ui import main as check_import_ui
        return check_import_ui()
    if cmd == 'importcheck':
        import unittest
        from tools.test_application_import import ApplicationImportTests
        suite = unittest.defaultTestLoader.loadTestsFromTestCase(ApplicationImportTests)
        return 0 if unittest.TextTestRunner(verbosity=2).run(suite).wasSuccessful() else 1
    if cmd == 'importcasecheck':
        from tools.check_application_import import main as check_import_cases
        return check_import_cases()
    if cmd == 'mapcasecheck':
        from tools.check_actual_boundary_maps import main as check_map_cases
        return check_map_cases()
    if cmd == 'mapcheck':
        import unittest
        from tools.test_boundary_map import BoundaryMapTests
        suite = unittest.defaultTestLoader.loadTestsFromTestCase(BoundaryMapTests)
        return 0 if unittest.TextTestRunner(verbosity=2).run(suite).wasSuccessful() else 1
    if cmd == "pdfcheck":
        from core.pdf_setup import main as check_pdf_setup
        return check_pdf_setup()
    if cmd == "hcad":
        from core import hcad_downloader
        sys.argv[0] = "hcad_downloader"
        hcad_downloader.main()
        return 0
    if cmd == "validate":
        from tools import validate
        sys.argv[0] = "validate"
        return int(validate.main() or 0)
    if cmd == "onboard":
        from tools import onboard_templates
        sys.argv[0] = "onboard_templates"
        return int(onboard_templates.main() or 0)
    print(f"Unknown command: {cmd}\n")
    print(USAGE)
    return 2


if __name__ == "__main__":
    sys.exit(main())
