"""Regression check: one notification per Step 1 property, even for shared mail."""
from __future__ import annotations
import argparse
import hashlib
import importlib
import json
import re
import tempfile
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import openpyxl
from core.paths import sample_input_dir, templates_dir
from core.templates import DEFAULT_ROLES

STEP1 = "Step 1 - Input Properties Info"
SAMPLES = {
    "mls": "MLS_Calculation_Sheet_Case_Info_v13_FILLED_SAMPLE.xlsx",
    "mbl": "MBL_Calculation_Sheet_Case_Info_v3_FILLED_SAMPLE.xlsx",
    "mls_renewal": "MLS_Renewal_Calculation_Sheet_Case_Info_v2_FILLED_SAMPLE.xlsx",
    "mbl_renewal": "MBL_Renewal_Calculation_Sheet_Case_Info_v2_FILLED_SAMPLE.xlsx",
}
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
OUT = None
PRIMARY_ROLES = {"mls": ("template",), "mbl": (),
                 "mls_renewal": ("sfr_template", "no_sfr_template"),
                 "mbl_renewal": ()}
PROPERTY_AREAS = ["6,101", "6,102", "6,103.5", "6,104"]

def doc_facts(path):
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    text = "\n".join("".join(p.itertext()) for p in root.findall(".//w:p", NS)).replace("\u00a0", " ")
    return {
        "property_value_colors": [r.find("w:rPr/w:color", NS).get("{" + NS["w"] + "}val")
                                 if r.find("w:rPr/w:color", NS) is not None else None
                                 for r in root.findall(".//w:r", NS)
                                 if any(area in "".join(r.itertext()) for area in PROPERTY_AREAS)],
        "property_lot_sizes": re.findall(r"your property's lot size is\s+([\d,]+(?:\.\d+)?) square feet", text),
        "review_minimums": re.findall(r"(?:minimum lot size is|proposed minimum lot size for this block would be)\s+([\d,]+)\s+square feet", text),
        "bold_numeric_runs": ["".join(r.itertext()).strip() for r in root.findall(".//w:r", NS)
                              if r.find("w:rPr/w:b", NS) is not None and
                              r.find("w:rPr/w:b", NS).get("{"+NS["w"]+"}val", "1") not in ("0", "false")],
        "legal_descriptions": re.findall(r"LT [1-4] BLK 1", text),
        "unfilled_property_fields": text.count("Property_Lot_Size"),
        "missing_property_area": text.count("[VERIFY PROPERTY LOT SIZE]"),
        "restrictions": re.findall(r"(?:smaller than|at least)\s+([\d,]+)\s+(square feet|feet)", text),
        "shared_owner": text.count("PROPERTY MAILOUT TEST OWNER"),
        "other_owner": text.count("SECOND MAILOUT TEST OWNER"),
        "mailing_address": text.count("123 SHARED MAILING AVENUE"),
        "sections": len(root.findall(".//w:sectPr", NS)),
        "drawings": len(root.findall(".//w:drawing", NS)),
        "footers": len(root.findall(".//w:footerReference", NS)),
    }

class PropertyNotificationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temporary = tempfile.TemporaryDirectory()
        cls.folder = Path(OUT or cls.temporary.name)
        cls.folder.mkdir(parents=True, exist_ok=True)
        cls.workbooks = {}
        cls.results = []
        for wizard, filename in SAMPLES.items():
            # Private copies only; cached sample values make the test independent
            # of desktop Excel and of any real case workbook.
            wb = openpyxl.load_workbook(sample_input_dir(wizard) / filename, data_only=True)
            ws = wb[STEP1]
            headers = [cell.value for cell in ws[1]]
            ws.delete_rows(2, ws.max_row)
            columns = {name: i + 1 for i, name in enumerate(headers) if name}
            for i in range(4):
                # Same owner and identical mailing address for three properties.
                # Two properties also share a site address, but have distinct accounts.
                # EXC remains a property requiring its own notification.
                values = {
                    "OWNER_MAILTO": "PROPERTY MAILOUT TEST OWNER" if i < 3 else "SECOND MAILOUT TEST OWNER",
                    "OWNER_LIST": "PROPERTY MAILOUT TEST OWNER" if i < 3 else "SECOND MAILOUT TEST OWNER",
                    "Mail_Addr_1": "123 SHARED MAILING AVENUE", "Mail_Addr_2": "",
                    "Mail_City": "HOUSTON", "Mail_State": "TX", "Mail_Zip": "77009",
                    "Site_addr_1": "101 TEST STREET" if i < 2 else f"{101 + i * 2} TEST STREET",
                    "TAX_ID": f"000000000000{i + 1}", "Total_Land_Area": 6103.5 if i == 2 else 6101 + i,
                    "group_dscr": "SINGLE FAMILY", "Land Use": "EXC" if i == 2 else "SFR",
                    "Property Description": f"TEST PROPERTY {i + 1}",
                    "Legal_Dscr_1": f"LT {i + 1} BLK 1", "Legal_Dscr_2": "MAILOUT TEST",
                }
                for key, value in values.items():
                    if key in columns:
                        ws.cell(i + 2, columns[key], value)
            # A styled empty row must not become a letter.
            ws.cell(7, 1).number_format = "@"
            dest = cls.folder / (wizard + ".xlsx")
            wb.save(dest)
            wb.close()
            cls.workbooks[wizard] = dest

    @classmethod
    def tearDownClass(cls):
        (cls.folder / "results.json").write_text(json.dumps(cls.results, indent=2), encoding="utf-8")
        cls.temporary.cleanup()

    def test_all_four_property_readers_preserve_shared_recipients(self):
        from wizards.mls.mailout_engine import read_owner_rows_from_step1
        for wizard, path in self.workbooks.items():
            with self.subTest(wizard=wizard):
                wb = openpyxl.load_workbook(path, data_only=True)
                if wizard == "mls":
                    rows = read_owner_rows_from_step1(wb)
                else:
                    rows = importlib.import_module("wizards." + wizard + ".wizard")._get_owners(wb)
                wb.close()
                self.assertEqual(4, len(rows))
                self.assertEqual(4, sum(r["Mail_Addr_1"] == "123 SHARED MAILING AVENUE" for r in rows))
                if wizard != "mls":
                    self.assertEqual([f"TEST PROPERTY {i}" for i in range(1, 5)],
                                     [r["Property_Description"] for r in rows])

    def check_letters(self, wizard, role):
        template_name = next(name for key, _label, name in DEFAULT_ROLES[wizard] if key == role)
        template = templates_dir(wizard) / template_name
        workbook = self.workbooks[wizard]
        before = hashlib.sha256(workbook.read_bytes()).hexdigest()
        map_name = "MLS_916_Map_SAMPLE.pdf" if wizard == "mls" else "sample_boundary_map.pdf"
        map_path = sample_input_dir(wizard) / map_name
        outputs = {}
        for preview in (True, False):
            dest = self.folder / f"{wizard}_{role}_{'preview' if preview else 'all'}.docx"
            if wizard == "mls":
                from wizards.mls.templates_2026 import generate_owner_letters
                count = generate_owner_letters(workbook, template, map_path, dest, preview=preview)
                self.assertEqual(1 if preview else 4, count)
            else:
                mod = importlib.import_module("wizards." + wizard + ".wizard")
                mod.generate_notification(workbook, template, map_path, dest,
                                          mode="preview" if preview else "all")
            outputs[preview] = doc_facts(dest)
        preview, all_letters = outputs[True], outputs[False]
        self.assertEqual(0, all_letters["unfilled_property_fields"])
        self.assertEqual(0, all_letters["missing_property_area"])
        if wizard in ("mls", "mls_renewal") and role in PRIMARY_ROLES[wizard]:
            if wizard == "mls":
                from wizards.mls.templates_2026 import build_context
                context = build_context(workbook)
            else:
                context = mod._get_renewal_context(workbook)
            minimum = str(context["Minimum_Lot_Size_Text"])
            self.assertNotIn(minimum, PROPERTY_AREAS)
            self.assertEqual([], all_letters["property_lot_sizes"])
            self.assertEqual([minimum], preview["review_minimums"])
            self.assertEqual([minimum] * 4, all_letters["review_minimums"])
            self.assertEqual([(minimum, "square feet")] * 4, all_letters["restrictions"])
            if wizard == "mls":
                self.assertEqual([f"LT {i} BLK 1" for i in range(1,5)], all_letters["legal_descriptions"])
                self.assertEqual(8, all_letters["bold_numeric_runs"].count(minimum))
        else:
            self.assertEqual([], all_letters["property_lot_sizes"])
        self.assertEqual(1, preview["shared_owner"])
        self.assertEqual(0, preview["other_owner"])
        self.assertEqual(3, all_letters["shared_owner"])
        self.assertEqual(1, all_letters["other_owner"])
        self.assertEqual(4, all_letters["mailing_address"])
        for key in ("sections", "drawings", "footers"):
            self.assertGreater(preview[key], 0, key)
            self.assertEqual(preview[key] * 4, all_letters[key], key)
        self.assertEqual(before, hashlib.sha256(workbook.read_bytes()).hexdigest())
        self.results.append({"wizard": wizard, "template": template_name, "properties": 4,
                             "letters": 4, "preview_letters": 1, "facts": all_letters})

    def test_mls_notification(self):
        self.check_letters("mls", "template")

    def test_missing_property_area_never_uses_case_minimum(self):
        from core.property_letters import property_lot_size, MISSING
        for invalid in (None, "", 0, -1, "NaN", "Infinity", "#VALUE!", "<Null>", "not known"):
            with self.subTest(area=invalid):
                self.assertEqual(MISSING, property_lot_size({"Property_Lot_Size": invalid,
                                                           "Minimum_Lot_Size": "6,250"}))
        self.assertEqual("6,103.5", property_lot_size({"Total_Land_Area": 6103.5}))
        self.assertEqual("6,250", property_lot_size({"Total_Land_Area": "6,250 square feet"}))

    def test_older_notification_wording_uses_property_area(self):
        from docx import Document
        from core.property_letters import fill_property_review
        prefixes = [
            "A preliminary review indicates the proposed minimum lot size for this block would be",
            "Preliminary staff review of the application indicates that the minimum lot size is",
        ]
        for prefix in prefixes:
            with self.subTest(prefix=prefix):
                doc = Document()
                p = doc.add_paragraph(prefix + " ")
                p.add_run("X, ").bold = True
                p.add_run("XXX").bold = True
                p.add_run(" feet." if "building line" in prefix else " square feet.")
                rule = doc.add_paragraph("If approved, lots could not be split into lots smaller than 6,250 square feet.")
                fill_property_review(doc, {"Total_Land_Area": "7,123.5"})
                self.assertIn("your property's lot size is 7,123.5 square feet", p.text)
                self.assertIn("6,250 square feet", rule.text)
                self.assertTrue(next(r for r in p.runs if "7,123.5" in r.text).bold)

    def test_split_property_field_and_missing_area_are_visible(self):
        from docx import Document
        from core.property_letters import fill_property_review, MISSING
        doc = Document()
        p = doc.add_paragraph("Your lot size: ")
        p.add_run("«Property_").bold = True
        p.add_run("Lot_Size»").bold = True
        fill_property_review(doc, {"Total_Land_Area": None, "Minimum_Lot_Size_Text": "6,250"})
        self.assertEqual("Your lot size: " + MISSING, p.text)

    def test_mbl_notification(self):
        self.check_letters("mbl", "template")

    def test_mls_renewal_sfr_notification(self):
        self.check_letters("mls_renewal", "sfr_template")

    def test_mls_renewal_no_sfr_notification(self):
        self.check_letters("mls_renewal", "no_sfr_template")

    def test_mbl_renewal_notification(self):
        self.check_letters("mbl_renewal", "owner_template")

    def test_related_owner_notices_keep_property_count(self):
        for wizard in SAMPLES:
            for role, _label, _filename in DEFAULT_ROLES[wizard]:
                if role in ("protest_template", "reschedule_template", "hearing_template"):
                    with self.subTest(wizard=wizard, role=role):
                        self.check_letters(wizard, role)

def main():
    global OUT
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", type=Path)
    args = parser.parse_args()
    OUT = args.out
    result = unittest.TextTestRunner(verbosity=2).run(
        unittest.defaultTestLoader.loadTestsFromTestCase(PropertyNotificationTests))
    return 0 if result.wasSuccessful() else 1

if __name__ == "__main__":
    raise SystemExit(main())
