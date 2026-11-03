"""Regression checks for the September MBL forms and optional RCA attachments."""
from contextlib import ExitStack
from pathlib import Path
import importlib, shutil, tempfile, unittest
from unittest.mock import patch
from docx import Document
from docx.oxml.ns import qn
import fitz
from core.paths import sample_input_dir
from core.templates import template_path
from core.mbl_templates_2026 import sfr_threshold

SAMPLES = {
    'mls': 'MLS_Calculation_Sheet_Case_Info_v13_FILLED_SAMPLE.xlsx',
    'mbl': 'MBL_Calculation_Sheet_Case_Info_v3_FILLED_SAMPLE.xlsx',
    'mls_renewal': 'MLS_Renewal_Calculation_Sheet_Case_Info_v2_FILLED_SAMPLE.xlsx',
    'mbl_renewal': 'MBL_Renewal_Calculation_Sheet_Case_Info_v2_FILLED_SAMPLE.xlsx',
}


def text(doc):
    return '\n'.join(''.join(p.xpath('.//w:t/text()')) for p in doc.element.xpath('.//w:p'))


class MBLTemplateTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
    def tearDown(self):
        self.temp.cleanup()

    def test_threshold_uses_counts_not_rounded_percentage(self):
        self.assertFalse(sfr_threshold(dict(Total_Lot_Count=101, SFR_Count=60, SFR_Percentage_Number='60')))
        self.assertTrue(sfr_threshold(dict(Total_Lot_Count=100, SFR_Count=60)))
        self.assertFalse(sfr_threshold(dict(Total_Lot_Count=10, SFR_Count=0)))
        for total, count in ((0,0), ('XX','X'), (10,11), ('NaN',2)):
            self.assertIsNone(sfr_threshold(dict(Total_Lot_Count=total,SFR_Count=count)))

    def test_director_new_wording_and_table_survive(self):
        m = importlib.import_module('wizards.mbl_renewal.wizard')
        wb = sample_input_dir('mbl_renewal') / SAMPLES['mbl_renewal']
        out = self.folder/'director.docx'
        m.generate_director_approval(wb,template_path('mbl_renewal','director_template'),out)
        doc=Document(out);t=text(doc)
        self.assertNotIn('{{',t)
        self.assertIn('Renewal No. 88REN',t)
        self.assertIn('building line of 35 feet',t)
        self.assertIn('Upon review by Planning Department staff',t)
        self.assertIn('Acknowledged and Reviewed:',t)
        self.assertIn('Misty Staunton',t)
        self.assertIn('Prepared by:',t)
        self.assertEqual(len(doc.tables[0].rows),6)
        self.assertIn('2400-2500',t)

    def test_routing_fields_wizard_and_approval_checkboxes(self):
        for key in ('mbl','mbl_renewal'):
            m=importlib.import_module(f'wizards.{key}.wizard')
            getter='_get_renewal_context' if key.endswith('renewal') else '_get_context'
            wb=sample_input_dir(key)/SAMPLES[key]
            for pc in (False,True):
                with self.subTest(wizard=key,pc=pc):
                    ctx=getattr(m,getter)(wb)
                    ctx.update(Council_District='D',Total_Lot_Count=100,SFR_Count=60,
                               Approval_Path='Planning Commission Approval' if pc else 'Director Approval')
                    out=self.folder/f'{key}_{pc}.docx'
                    with patch.object(m,getter,return_value=ctx):
                        m.generate_routing_form(wb,template_path(key,'routing_template'),out)
                    d=Document(out);t=text(d)
                    self.assertNotIn('{{',t)
                    self.assertFalse(any(n.get(qn('w:val'),'').upper() in ('EE0000','FF0000','C00000')
                                         for n in d.element.iter(qn('w:color'))))
                    self.assertIn('Council District: D',t)
                    self.assertIn('Yes ☒ No ☐',t)
                    self.assertIn(('☒' if pc else '☐')+' Planning Commission Approval',t)
                    self.assertIn(('☐' if pc else '☒')+" Director's Approval",t)
                    self.assertIn(('☒' if key.endswith('renewal') else '☐')+' SMBL Block Renewal',t)
                    self.assertIn(('☒' if key.endswith('renewal') else '☐')+' Original Ordinance',t)
                    self.assertNotIn('Tax Records',t)
                    self.assertNotIn('Copy of Subdivision Plat',t)
                    if key=='mbl':self.assertIn('Previous ordinance: N/A',t)
                    else:self.assertIn('Building Line Block Renewal Application No.',t)
                    self.assertEqual(len(d.sections),1)

    def test_current_packet_builders_optional_attachment_matrix(self):
        marker=self.folder/'provided.pdf'
        with fitz.open() as pdf:
            pdf.new_page().insert_text((72,72),'Provided attachment');pdf.save(marker)
        for key in SAMPLES:
            m=importlib.import_module('wizards.mls.templates_2026' if key=='mls' else f'wizards.{key}.wizard')
            for plat, hcad in ((False,False),(False,True),(True,False),(True,True)):
                for pc in (False,True):
                    with self.subTest(key=key,plat=plat,hcad=hcad,pc=pc), ExitStack() as stack:
                        wb=sample_input_dir(key)/SAMPLES[key]
                        getter='build_context' if key=='mls' else ('_get_renewal_context' if key.endswith('renewal') else '_get_context')
                        ctx=getattr(m,getter)(wb)
                        ctx['Approval_Path']='Planning Commission Approval' if pc else 'Director Approval'
                        stack.enter_context(patch.object(m,getter,return_value=ctx))
                        calls=[]
                        def convert(path,work):
                            self.assertTrue(str(path).strip())
                            calls.append(str(path));return marker
                        stack.enter_context(patch.object(m.ME if key=='mls' else m,'ensure_pdf' if key=='mls' else '_ensure_pdf',side_effect=convert))
                        args=dict(templates={k:template_path(key,k+'_template') for k in ('routing','director','pc')},
                                  map_path=marker,tax_map_path=marker,output_pdf=self.folder/'packet.pdf',
                                  subdivision_plat_path=marker if plat else '',include_staff_analysis=False)
                        if key=='mls':args.update(excel_path=wb,hcad_records=[marker] if hcad else [])
                        else:args.update(workbook_path=wb,hcad_records_path=str(marker) if hcad else '')
                        if key.endswith('renewal'):args['original_ordinance_path']=marker
                        m.generate_rca_packet_pdf(**args)
                        with fitz.open(args['output_pdf']) as pdf:
                            self.assertEqual(len(pdf),4+int(key.endswith('renewal'))+int(plat)+int(hcad))

    def test_legacy_packet_builders_accept_empty_and_none_attachments(self):
        for key in ('mls','mbl'):
            m=importlib.import_module(f'wizards.{key}.'+('mailout_engine' if key=='mls' else 'mbl_mailout_engine'))
            for missing in ('',None):
                with self.subTest(key=key,missing=missing),ExitStack() as stack:
                    marker=self.folder/'provided.pdf'
                    with fitz.open() as pdf:
                        pdf.new_page().insert_text((72,72),'Required attachment');pdf.save(marker)
                    stack.enter_context(patch.object(m,'ensure_pdf',return_value=marker))
                    for generator in ('generate_routing_form','generate_director_approval'):
                        stack.enter_context(patch.object(m,generator))
                    export='export_step2_staff_analysis_pdf' if key=='mls' else 'export_mbl_step2_staff_analysis_pdf'
                    stack.enter_context(patch.object(m,export,side_effect=lambda wb,out:shutil.copyfile(marker,out)))
                    wb=sample_input_dir(key)/SAMPLES[key]
                    out=self.folder/'legacy.pdf'
                    fn=m.generate_rca_packet_pdf if key=='mls' else m.generate_mbl_packet_pdf
                    fn(wb,dict(routing='unused',director='unused',pc='unused'),marker,missing,marker,missing,out)
                    with fitz.open(out) as pdf:self.assertGreaterEqual(len(pdf),4)


def main():
    result=unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(MBLTemplateTests))
    return 0 if result.wasSuccessful() else 1

if __name__=='__main__':raise SystemExit(main())
