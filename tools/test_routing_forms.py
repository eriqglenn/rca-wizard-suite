"""Routing form layout, content preservation, and packet pagination checks."""
from __future__ import annotations

import argparse
from collections import Counter
from contextlib import ExitStack
from copy import deepcopy
import hashlib
import importlib
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from docx import Document
from docx.oxml.ns import qn
import fitz

from core.paths import sample_input_dir, templates_dir
from core.routing_form import prepare_routing_form, routing_pdf_one_page
from core.templates import DEFAULT_ROLES

OUT = None
ACTUAL = None
ENGINE = 'none'
SAMPLES = {
    'mls': 'MLS_Calculation_Sheet_Case_Info_v13_FILLED_SAMPLE.xlsx',
    'mbl': 'MBL_Calculation_Sheet_Case_Info_v3_FILLED_SAMPLE.xlsx',
    'mls_renewal': 'MLS_Renewal_Calculation_Sheet_Case_Info_v2_FILLED_SAMPLE.xlsx',
    'mbl_renewal': 'MBL_Renewal_Calculation_Sheet_Case_Info_v2_FILLED_SAMPLE.xlsx',
}
CHECKED = '{http://schemas.microsoft.com/office/word/2010/wordml}checked'


def document_content(doc):
    root = doc.element
    words = Counter(' '.join(t.text or '' for t in root.iter(qn('w:t'))).split())
    checks = [dict(node.attrib) for node in root.iter(CHECKED)]
    # Compare the formatted text runs, including checkbox glyphs inside SDTs.
    runs = []
    for run in root.iter(qn('w:r')):
        text = ''.join(t.text or '' for t in run.iter(qn('w:t')))
        if text.strip():
            props = run.find(qn('w:rPr'))
            formatting = () if props is None else tuple(
                (node.tag, tuple(sorted(node.attrib.items()))) for node in props.iter())
            runs.append((text, formatting))
    return words, checks, Counter(runs)


class RoutingFormTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        cls.folder = Path(OUT or cls.temp.name).resolve()
        cls.folder.mkdir(parents=True, exist_ok=True)
        cls.results = []
        cls.engines = {}
        cls.contexts = {}
        for key in SAMPLES:
            module = 'wizards.mls.templates_2026' if key == 'mls' else f'wizards.{key}.wizard'
            engine = cls.engines[key] = importlib.import_module(module)
            getter = cls.getter(key)
            cls.contexts[key] = getattr(engine, getter)(sample_input_dir(key) / SAMPLES[key])
        cls.overflow = cls.folder / 'two_page_routing_fixture.pdf'
        with fitz.open() as doc:
            doc.new_page().insert_text((72, 72), 'Legal Department Routing Form - first page')
            doc.new_page().insert_text((72, 72), 'Tax Records\nOriginal Ordinance')
            doc.save(cls.overflow)
        cls.attachment = cls.folder / 'attachment_fixture.pdf'
        with fitz.open() as doc:
            doc.new_page().insert_text((72, 72), 'APPROVAL / ATTACHMENT')
            doc.save(cls.attachment)

    @classmethod
    def tearDownClass(cls):
        (cls.folder / 'results.json').write_text(json.dumps(cls.results, indent=2), encoding='utf-8')
        cls.temp.cleanup()

    @staticmethod
    def getter(key):
        return 'build_context' if key == 'mls' else ('_get_renewal_context' if 'renewal' in key else '_get_context')

    @staticmethod
    def template(key):
        return templates_dir(key) / next(role[2] for role in DEFAULT_ROLES[key] if role[0] == 'routing_template')

    def assert_layout(self, doc):
        self.assertEqual(len(list(doc.element.body.iter(qn('w:sectPr')))), 1)
        self.assertEqual(len(doc.tables), 2)
        for table in doc.tables:
            self.assertEqual(len(table.rows), 1)
            self.assertIsNotNone(table.rows[0]._tr.trPr.find(qn('w:cantSplit')))
        before = document_content(doc)
        paragraphs = len(list(doc.element.iter(qn('w:p'))))
        prepare_routing_form(doc)
        self.assertEqual(document_content(doc), before)
        self.assertEqual(len(list(doc.element.iter(qn('w:p')))), paragraphs)
        self.assertEqual(len(doc.tables), 2)

    def render(self, path, last_items):
        if ENGINE == 'none':
            return
        from core.office_convert import office_to_pdf
        for engine in ('libreoffice', 'msoffice') if ENGINE == 'both' else (ENGINE,):
            folder = self.folder / engine / path.stem
            folder.mkdir(parents=True, exist_ok=True)
            with patch.dict(os.environ, {'RCA_OFFICE_ENGINE': engine}):
                pdf = Path(office_to_pdf(path, folder))
            with fitz.open(pdf) as doc:
                self.assertEqual(len(doc), 1, f'{path.name} in {engine}')
                text = ' '.join(doc[0].get_text().split())
                for item in last_items:
                    self.assertIn(item, text)
                doc[0].get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False).save(folder / 'page-1.png')
            self.results.append({'docx': str(path), 'engine': engine, 'pages': 1, 'pdf': str(pdf)})

    def test_all_four_generators_preserve_text_formatting_and_checkboxes(self):
        for key, engine in self.engines.items():
            for approval in ('Director Approval', 'Planning Commission Approval'):
                with self.subTest(wizard=key, approval=approval):
                    ctx = deepcopy(self.contexts[key])
                    ctx['Approval_Path'] = approval
                    name = key + ('_director' if approval == 'Director Approval' else '_pc')
                    old_path, path = self.folder / (name + '_before.docx'), self.folder / (name + '.docx')
                    workbook = sample_input_dir(key) / SAMPLES[key]
                    digest = hashlib.sha256(workbook.read_bytes()).digest()
                    with patch.object(engine, self.getter(key), return_value=ctx):
                        with patch.object(engine, 'prepare_routing_form', lambda doc: None):
                            engine.generate_routing_form(workbook, self.template(key), old_path)
                        engine.generate_routing_form(workbook, self.template(key), path)
                    self.assertEqual(digest, hashlib.sha256(workbook.read_bytes()).digest())
                    before, after = Document(old_path), Document(path)
                    self.assertEqual(document_content(before), document_content(after))
                    self.assert_layout(after)
                    if approval == 'Director Approval':
                        last = ['Tax Map'] + (['Original Ordinance'] if key == 'mbl_renewal' else [])
                        self.render(path, last)

    def test_long_location_stays_one_page(self):
        for key, engine in self.engines.items():
            with self.subTest(wizard=key):
                ctx = deepcopy(self.contexts[key])
                ctx['Area_Description'] = ('100-200 blocks of Martin Luther King Junior Boulevard, north and south sides, '
                    'between West Little York Road and South Sam Houston Parkway East, including the adjoining '
                    'properties identified in the application boundary')
                ctx['Planner_Name'] = 'Alexandra Catherine Planning Reviewer'
                path = self.folder / (key + '_long_location.docx')
                with patch.object(engine, self.getter(key), return_value=ctx):
                    engine.generate_routing_form(sample_input_dir(key) / SAMPLES[key], self.template(key), path)
                self.assert_layout(Document(path))
                self.render(path, ['Attachments Provided:'])

    def test_converter_overflow_is_retained_on_one_page(self):
        before = self.overflow.read_bytes()
        result = routing_pdf_one_page(self.overflow, self.folder)
        with fitz.open(result) as doc:
            self.assertEqual(len(doc), 1)
            for marker in ('Legal Department Routing Form', 'Tax Records', 'Original Ordinance'):
                self.assertIn(marker, doc[0].get_text())
        self.assertEqual(self.overflow.read_bytes(), before)
        before = self.attachment.read_bytes()
        self.assertEqual(routing_pdf_one_page(self.attachment, self.folder), self.attachment)
        self.assertEqual(self.attachment.read_bytes(), before)

    def test_all_four_packet_builders_keep_overflow_on_routing_page(self):
        # Force the converter to spill, then exercise each actual PDF merger.
        def convert(path, work):
            return self.overflow if 'Routing' in Path(path).name else self.attachment

        for key, engine in self.engines.items():
            for approval in ('Director Approval', 'Planning Commission Approval'):
                with self.subTest(wizard=key, approval=approval), ExitStack() as stack:
                    ctx = deepcopy(self.contexts[key])
                    ctx['Approval_Path'] = approval
                    stack.enter_context(patch.object(engine, self.getter(key), return_value=ctx))
                    stack.enter_context(patch.object(engine, 'generate_routing_form'))
                    stack.enter_context(patch.object(engine, 'generate_director_approval'))
                    if key == 'mls':
                        stack.enter_context(patch.object(engine.ME, 'ensure_pdf', side_effect=convert))
                        stack.enter_context(patch.object(engine.ME, 'generate_planning_commission_approval'))
                    else:
                        stack.enter_context(patch.object(engine, '_ensure_pdf', side_effect=convert))
                        stack.enter_context(patch.object(engine, 'generate_pc_approval'))
                    templates = {k: str(self.template(key)) for k in ('routing', 'director', 'pc')}
                    path = self.folder / (key + '_' + approval.split()[0] + '_packet.pdf')
                    args = {'templates': templates, 'map_path': self.attachment, 'tax_map_path': self.attachment,
                            'output_pdf': path, 'include_staff_analysis': False}
                    if key == 'mls':
                        args['excel_path'] = sample_input_dir(key) / SAMPLES[key]
                    else:
                        args.update(workbook_path=sample_input_dir(key) / SAMPLES[key], subdivision_plat_path='', hcad_records_path='')
                    if 'renewal' in key:
                        args['original_ordinance_path'] = self.attachment
                    engine.generate_rca_packet_pdf(**args)
                    with fitz.open(path) as doc:
                        self.assertEqual(len(doc), 5 if 'renewal' in key else 4)
                        self.assertIn('Legal Department Routing Form', doc[0].get_text())
                        self.assertIn('Tax Records', doc[0].get_text())
                        self.assertIn('Original Ordinance', doc[0].get_text())
                        for page in list(doc)[1:]:
                            self.assertIn('APPROVAL / ATTACHMENT', page.get_text())
                    self.results.append({'packet': str(path), 'wizard': key, 'approval': approval, 'routing_pages': 1})

    def test_actual_mbl_renewal_workbook(self):
        if not ACTUAL:
            self.skipTest('No actual renewal workbook supplied')
        workbook = Path(ACTUAL)
        digest = hashlib.sha256(workbook.read_bytes()).digest()
        engine = self.engines['mbl_renewal']
        path = self.folder / 'actual_MBL_renewal_Routing_Form.docx'
        engine.generate_routing_form(workbook, self.template('mbl_renewal'), path)
        self.assertEqual(digest, hashlib.sha256(workbook.read_bytes()).digest())
        doc = Document(path)
        self.assert_layout(doc)
        text = ' '.join(''.join(t.text or '' for t in p.iter(qn('w:t')))
                        for p in doc.element.iter(qn('w:p')))
        self.assertIn('Tax Map', text)
        self.assertIn('Original Ordinance', text)
        self.render(path, ['Tax Map', 'Original Ordinance'])


def main():
    global OUT, ACTUAL, ENGINE
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--out')
    parser.add_argument('--workbook', help='Optional actual MBL renewal calculation sheet (read only)')
    parser.add_argument('--engine', choices=('none', 'libreoffice', 'msoffice', 'both'), default='none')
    args = parser.parse_args()
    OUT, ACTUAL, ENGINE = args.out, args.workbook, args.engine
    result = unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(RoutingFormTests))
    return 0 if result.wasSuccessful() else 1


if __name__ == '__main__':
    raise SystemExit(main())
