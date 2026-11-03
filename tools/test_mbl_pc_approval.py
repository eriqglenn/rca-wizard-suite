"""Checks for the supplied 2026 MBL Renewal PC approval form."""
from pathlib import Path
import argparse, copy, tempfile, unittest
from unittest.mock import patch
from docx import Document
from docx.oxml.ns import qn
from core.paths import sample_input_dir
from core.templates import template_path
from wizards.mbl_renewal import wizard as m

OUTPUT = None


class PCApprovalTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.workbook = sample_input_dir('mbl_renewal') / 'MBL_Renewal_Calculation_Sheet_Case_Info_v2_FILLED_SAMPLE.xlsx'
        cls.template = template_path('mbl_renewal', 'pc_template')
        cls.context = m._get_renewal_context(cls.workbook)
        cls.tmp = tempfile.TemporaryDirectory()
        cls.output = OUTPUT or Path(cls.tmp.name)
        cls.output.mkdir(parents=True, exist_ok=True)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def generate(self, name, **changes):
        ctx = dict(self.context, **changes)
        path = self.output / (name + '.docx')
        with patch.object(m, '_get_renewal_context', return_value=ctx):
            m.generate_pc_approval(self.workbook, self.template, path)
        doc = Document(path)
        text = '\n'.join(p.text for p in m._iter_all_paragraphs(doc))
        self.assertNotIn('{{', text)
        return doc, text

    def test_new_wording_and_independent_counts(self):
        doc, text = self.generate('completed_approval', PC_Noncorner_Lot_Count=9, PC_Approval_Date='September 24, 2026')
        self.assertIn('Renewal No. 88REN', text)
        self.assertIn('seventeen (17) lots. Nine (9) lots (excluding corner lots) have a building line of 35 feet.', text)
        self.assertIn('76% of the lots', text)
        self.assertIn('Over 60%', text)
        self.assertIn('Lisa M. Clark, Chair', text)
        self.assertIn('Sonny Garza, Vice-Chair', text)
        self.assertEqual(text.count('September 24, 2026'), 2)
        self.assertNotIn('Less than 25 percent', text)
        self.assertNotIn('ABC Street', text)
        # Finding text remains bold while legal criterion remains italic.
        for p in m._iter_all_paragraphs(doc):
            if 'The application is for the' in p.text:
                findings = [r for r in p.runs if '2400-2500' in r.text]
                self.assertEqual(len(findings), 1)
                self.assertTrue(findings[0].bold)
                self.assertTrue(any(r.italic for r in p.runs if 'renewal' in r.text))

    def test_blank_count_is_not_replaced_with_compliant_or_total(self):
        _, text = self.generate('blank_count', PC_Noncorner_Lot_Count='')
        self.assertIn('[ENTER COUNT] lots (excluding corner lots)', text)
        self.assertEqual(sum(p.text == 'Date' for p in m._iter_all_paragraphs(Document(self.output/'blank_count.docx'))), 2)

    def test_zero_is_a_valid_planner_count(self):
        _, text = self.generate('zero_count', PC_Noncorner_Lot_Count=0)
        self.assertIn('Zero (0) lots (excluding corner lots)', text)

    def test_invalid_counts_remain_visible(self):
        for raw in ('-1', '1.5', 'eighteen', '18', 'nan', 'inf'):
            with self.subTest(raw=raw):
                fields=m._pc_approval_fields(dict(self.context, PC_Noncorner_Lot_Count=raw))
                self.assertEqual(fields['{{MBL_REN_NONCORNER_LOTS}}'], '[ENTER COUNT]')

    def test_protest_grammar(self):
        for phrase in ('No protest was received.', 'One (1) protest was filed.', 'Three (3) protests were filed.'):
            with self.subTest(phrase=phrase):
                _, text=self.generate('protest', Protest_Sentence=phrase, PC_Noncorner_Lot_Count=7)
                self.assertIn(phrase, text)

    def test_sfr_is_lot_percentage_not_support_or_area(self):
        _, text=self.generate('percentage', SFR_Percentage_Number='72', Support_Percentage='94', PC_Noncorner_Lot_Count=8)
        self.assertIn('72% of the lots', text)
        self.assertNotIn('94%', text)

    def test_geometry_and_table_properties_preserved(self):
        before=Document(self.template)
        after,_=self.generate('layout', PC_Noncorner_Lot_Count=9)
        for doc in (before,after):
            self.assertEqual(len(doc.sections), 1)
        self.assertEqual(before.sections[0]._sectPr.xml, after.sections[0]._sectPr.xml)
        for xpath in ('.//w:tblPr','.//w:tblGrid','.//w:tcPr','.//w:trPr','.//w:pPr'):
            self.assertEqual([e.xml for e in before.element.xpath(xpath)], [e.xml for e in after.element.xpath(xpath)])

    def test_real_workbook_default_path_and_field(self):
        self.assertEqual(self.context.get('PC_Noncorner_Lot_Count'), '')
        dest=self.output/'from_default_workbook.docx'
        self.assertEqual(m.generate_pc_approval(self.workbook,self.template,dest),1)
        text='\n'.join(p.text for p in m._iter_all_paragraphs(Document(dest)))
        self.assertIn('[ENTER COUNT]',text)
        self.assertNotIn('{{',text)


def main():
    global OUTPUT
    parser=argparse.ArgumentParser()
    parser.add_argument('--out',type=Path)
    args=parser.parse_args()
    OUTPUT=args.out
    result=unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(PCApprovalTests))
    return 0 if result.wasSuccessful() else 1


if __name__ == '__main__':
    raise SystemExit(main())
