"""Qualifying counts and all four draft ordinance generators, offline."""
from __future__ import annotations
import argparse
import hashlib
import importlib
import json
import tempfile
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import openpyxl
from core.paths import sample_input_dir, templates_dir
from core.templates import DEFAULT_ROLES
from wizards.mls import mailout_engine as ME, templates_2026 as MLS

OUT = None
ACTUAL = None
MAP = None
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
SAMPLES = {
    'mls': 'MLS_Calculation_Sheet_Case_Info_v13_FILLED_SAMPLE.xlsx',
    'mbl': 'MBL_Calculation_Sheet_Case_Info_v3_FILLED_SAMPLE.xlsx',
    'mls_renewal': 'MLS_Renewal_Calculation_Sheet_Case_Info_v2_FILLED_SAMPLE.xlsx',
    'mbl_renewal': 'MBL_Renewal_Calculation_Sheet_Case_Info_v2_FILLED_SAMPLE.xlsx',
}

def text_from_doc(path):
    with ZipFile(path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
    return ' '.join(''.join(p.itertext()) for p in root.findall('.//w:p', NS)).replace('\u00a0', ' ')

class OrdinanceCountTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        cls.folder = Path(OUT or cls.temp.name)
        cls.folder.mkdir(parents=True, exist_ok=True)
        cls.results = []
        import fitz
        cls.map = cls.folder / 'test_map.pdf'
        with fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((72, 72), 'TEST FIXTURE - ordinance map placement')
            doc.save(cls.map)

    @classmethod
    def tearDownClass(cls):
        (cls.folder / 'results.json').write_text(json.dumps(cls.results, indent=2), encoding='utf-8')
        cls.temp.cleanup()

    def workbook(self, rows, minimum=5000, manual=None, total=None):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = ME.STEP2_SHEET
        ws['E4'] = f'{minimum}-Square Foot' if minimum is not None else None
        eligible = 0
        for i, (land, area) in enumerate(rows, 17):
            ws.cell(i, 5, i - 16)  # rank index persists across empty template rows
            if land is None and area is None:
                continue
            ws.cell(i, 1, '101 SAME ADDRESS')
            ws.cell(i, 2, land)
            ws.cell(i, 4, area)
            ws.cell(i, 5, i - 16)
            eligible += land != 'EXC'
        ws['A15'] = eligible if total is None else total
        if manual is not None:
            ci = wb.create_sheet(ME.CASE_INFO_SHEET)
            ci.append(['Lots_Meeting_Minimum', manual])
        path = self.folder / 'unit.xlsx'
        wb.save(path)
        wb.close()
        return path

    def count(self, rows, **kwargs):
        path = self.workbook(rows, **kwargs)
        before = hashlib.sha256(path.read_bytes()).hexdigest()
        value = ME._read_step2_official_qualifying_values(path)
        self.assertEqual(before, hashlib.sha256(path.read_bytes()).hexdigest())
        return value

    def test_ties_and_duplicate_addresses_count_each_property(self):
        # Crossing 70% of area happens at 11; ALL 14 qualifying lots count.
        result = self.count([('SFR', 5000)] * 14 + [('VAC', 1000)] * 6)
        self.assertEqual(result['Qualifying_Lot_Count'], '14')
        self.assertEqual(result['Qualifying_Lot_Count_Words'], 'fourteen')

    def test_exclusions_zero_area_decimal_and_blank_rows(self):
        result = self.count([('SFR', 5000), ('VAC', '5,000.50'), ('MF', 4999.99),
                             ('EXC', 9000), (None, None), ('COM', 0)])
        self.assertEqual(result['Qualifying_Lot_Count'], '2')

    def test_missing_case_info_or_blank_manual_automates(self):
        for manual in (None, '', '   '):
            with self.subTest(manual=manual):
                self.assertEqual(self.count([('SFR', 6000)], manual=manual)['Qualifying_Lot_Count'], '1')

    def test_manual_count_preserved_including_zero(self):
        for manual, expected in ((14, '14'), ('14.0', '14'), (0, '0'), ('51', '51')):
            with self.subTest(manual=manual):
                result = self.count([('SFR', 6000)], manual=manual)
                self.assertEqual(result['Qualifying_Lot_Count'], expected)

    def test_invalid_manual_is_not_silently_rounded(self):
        for manual in ('fourteen', -1, '14.5', 'NaN', 'Infinity'):
            with self.subTest(manual=manual):
                self.assertEqual(self.count([('SFR', 6000)], manual=manual)['Qualifying_Lot_Count'], 'X')

    def test_incomplete_data_retains_placeholder(self):
        for rows, kwargs in (([('SFR', None)], {}), ([('SFR', '#NUM!')], {}),
                             ([('SFR', 6000)], {'minimum': None}),
                             ([(None, 6000)], {}), ([], {})):
            with self.subTest(rows=rows, kwargs=kwargs):
                self.assertEqual(self.count(rows, **kwargs)['Qualifying_Lot_Count'], 'X')

    def test_large_counts_are_words(self):
        result = self.count([('SFR', 6000)] * 109)
        self.assertEqual(result['Qualifying_Lot_Count_Words'], 'one hundred nine')
        from wizards.mbl import mbl_mailout_engine as MBL
        for count, words in ((0, 'zero'), (14, 'fourteen'), (31, 'thirty-one'), (40, 'forty'), (84, 'eighty-four')):
            self.assertEqual(MBL._num_words(count), words)

    def mls_fixture(self):
        source = sample_input_dir('mls') / SAMPLES['mls']
        wb = openpyxl.load_workbook(source, data_only=True)
        ws = wb[ME.STEP2_SHEET]
        for row in ws.iter_rows(min_row=17):
            for cell in row:
                cell.value = None
        areas = [5000] * 14 + [1000] * 6
        for r, area in enumerate(areas, 17):
            for c, val in ((1, f'{100 + r} TEST STREET'), (2, 'SFR'), (4, area), (5, r-16)):
                ws.cell(r, c, val)
        ws['E4'] = '5000-Square Foot'
        ws['K16'] = 5000
        ws['A15'] = 20
        ws['F11'] = 20
        ws['H11'] = 1
        ws['H8'] = 0.66
        for row in wb[ME.CASE_INFO_SHEET].iter_rows():
            if row[0].value == 'Lots_Meeting_Minimum':
                row[1].value = None
        # Matching exhibit rows, with separate properties even when owners repeat.
        ws1 = wb[ME.STEP1_SHEET]
        values = [c.value for c in ws1[2]]
        headers = {c.value: c.column for c in ws1[1]}
        ws1.delete_rows(2, ws1.max_row)
        for i, area in enumerate(areas, 2):
            for col, val in enumerate(values, 1):
                ws1.cell(i, col, val)
            for key, val in {'Site_addr_1': f'{115+i} TEST STREET', 'TAX_ID': f'{i:013d}',
                             'Total_Land_Area': area, 'Lot Size': area,
                             'Legal_Dscr_1': f'LT {i-1} BLK 1', 'Property Description': f'{115+i} (LT {i-1} BLK 1)'}.items():
                if key in headers:
                    ws1.cell(i, headers[key], val)
        path = self.folder / 'MLS_fourteen_of_twenty.xlsx'
        wb.save(path)
        wb.close()
        return path

    def test_context_zero_and_shared_document_counts(self):
        path = self.mls_fixture()
        for build in (ME.build_ordinance_context, ME.build_director_context, ME.build_pc_context, ME.build_routing_context):
            self.assertEqual(build(path)['Qualifying_Lot_Count'], '14')
        wb = openpyxl.load_workbook(path)
        wb[ME.STEP2_SHEET]['E4'] = '9000-Square Foot'
        wb.save(path)
        wb.close()
        self.assertEqual(MLS.build_context(path)['Qualifying_Phrase'], 'zero (0) out of twenty (20)')

    def test_all_four_draft_ordinances(self):
        engines = {'mls': MLS, 'mbl': importlib.import_module('wizards.mbl.mbl_mailout_engine'),
                   'mls_renewal': importlib.import_module('wizards.mls_renewal.wizard'),
                   'mbl_renewal': importlib.import_module('wizards.mbl_renewal.wizard')}
        for key, engine in engines.items():
            with self.subTest(wizard=key):
                workbook = self.mls_fixture() if key == 'mls' else sample_input_dir(key) / SAMPLES[key]
                template = templates_dir(key) / next(x[2] for x in DEFAULT_ROLES[key] if x[0] == 'ordinance_template')
                output = self.folder / (key + '_Draft_Ordinance.docx')
                before = hashlib.sha256(workbook.read_bytes()).hexdigest()
                engine.generate_draft_ordinance(workbook, template, self.map, output)
                text = text_from_doc(output)
                self.assertNotIn('X (X)', text)
                self.assertNotIn('{{Qualifying', text)
                if key == 'mls':
                    self.assertIn('fourteen (14) out of twenty (20)', text)
                elif key == 'mbl':
                    context = engine._mbl_build_ordinance_context(workbook)
                    phrase = f"{context['Qualifying_Lot_Count_Words']} ({context['Qualifying_Lot_Count']}) out of {context['Total_Lot_Count_Words']} ({context['Total_Lot_Count']})"
                    self.assertIn(phrase, text)
                else:
                    self.assertIn('recommends renewal', text)
                    self.assertNotIn('out of', text)
                self.assertEqual(before, hashlib.sha256(workbook.read_bytes()).hexdigest())
                self.results.append({'wizard': key, 'output': str(output), 'source_unchanged': True})

    def test_actual_edfield_case(self):
        if not ACTUAL:
            self.skipTest('optional real case workbook')
        path = Path(ACTUAL)
        before = hashlib.sha256(path.read_bytes()).hexdigest()
        ctx = MLS.build_context(path)
        self.assertEqual(ctx['Qualifying_Phrase'], 'thirteen (13) out of twenty (20)')
        self.assertEqual(ctx['Minimum_Lot_Size_Text'], '6,380')
        output = self.folder / 'MLS_905_Draft_Ordinance.docx'
        MLS.generate_draft_ordinance(path, templates_dir('mls') / 'MLS_Ordinance_Template.docx', MAP or '', output)
        self.assertIn(ctx['Qualifying_Phrase'], text_from_doc(output))
        self.assertEqual(before, hashlib.sha256(path.read_bytes()).hexdigest())
        self.results.append({'wizard': 'mls', 'case': 'MLS 905', 'phrase': ctx['Qualifying_Phrase'], 'output': str(output), 'source_unchanged': True})

def main():
    global OUT, ACTUAL, MAP
    parser = argparse.ArgumentParser()
    parser.add_argument('--out')
    parser.add_argument('--workbook')
    parser.add_argument('--map')
    args = parser.parse_args()
    OUT, ACTUAL, MAP = args.out, args.workbook, args.map
    suite = unittest.defaultTestLoader.loadTestsFromTestCase(OrdinanceCountTests)
    return 0 if unittest.TextTestRunner(verbosity=2).run(suite).wasSuccessful() else 1

if __name__ == '__main__':
    raise SystemExit(main())
