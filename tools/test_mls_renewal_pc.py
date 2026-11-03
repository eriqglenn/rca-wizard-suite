"""MLS Renewal PC source mapping, map symbols, missing facts, and UI regression."""
import io, tempfile, unittest, copy
from pathlib import Path
from unittest.mock import patch
from openpyxl import Workbook, load_workbook
from docx import Document
from pypdf import PdfReader
from wizards.mls_renewal.pc_data import read_pc_facts
from wizards.mls_renewal.pc_report import report_context
from wizards.mls_renewal.pc_speech import speech_context, defaults
from wizards.mls.pc_speech import fill_speech, generate_speech
from wizards.mls.pc_report import fill_template, generate_pc_report
from wizards.mls.pc_data import attach_pc_rows
from wizards.mls.pc_map import render_pc_map, pc_defaults
from core.paths import sample_input_dir


def case_fixture(folder):
    p=Path(folder)/'renewal.xlsx'; wb=Workbook();s1=wb.active;s1.title='Step 1'
    s1.append(['Property Description','group_dscr'])
    s=wb.create_sheet('Step 2');s['A1']='SPECIAL MINIMUM LOT SIZE BLOCK RENEWAL'
    for cell,value in {'B2':'555REN','C2':'Lot Size to Renew','D2':5000,'B4':'Example Street','D4':'1700 block of Example Street','B5':'West Street','D5':'East Street','B6':'north','D6':'south'}.items():s[cell]=value
    for col,value in enumerate(['Property Description','Land Use','Corner Lot?','Original Lot Size','Current Lot Size','Meets Minimum'],1):s.cell(14,col,value)
    # Same legal parent property split into three equal lots; EXC stays outside totals.
    for i,(area,lu) in enumerate([(5000,'SFR'),(5000,'SFR'),(5000,'SFR'),(1000,'COM'),(9000,'EXC')],15):
        values=[f'{1700+i} (LT {i-14} BLK 1)',lu,'N',15000,area,'YES']
        for c,value in enumerate(values,1):s.cell(i,c,value)
        s1.append([values[0],'Single-Family Residential' if lu=='SFR' else 'Commercial'])
    ci=wb.create_sheet('Case Info')
    for pair in [('Original_Ordinance_No','2006-121'),('Original_Passage_Date','March 1, 2006'),('Original_Effective_Date','March 15, 2006'),('PC_Agenda_Number','12'),('PC_Hearing_Date','October 8, 2026'),('Protest_Count',1)]:ci.append(pair)
    wb.save(p);wb.close();return p


class RenewalPCTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup)
        self.folder=Path(self.temp.name);self.path=case_fixture(self.folder);self.facts=read_pc_facts(self.path)

    def test_current_areas_split_lots_and_flags_are_not_support(self):
        f=self.facts
        self.assertEqual((f['eligible_count'],f['sfr_count'],f['excluded_count']),(4,3,1))
        self.assertEqual((f['qualifying_count'],f['sfr_lot_percent'],f['sfr_area_percent']),(3,'75','94'))
        self.assertEqual(f['minimum'],5000);self.assertIsNone(f['signed_count'])
        self.assertFalse(any(r['signed'] for r in f['rows']))
        self.assertEqual(f['variance']['percent'],'0')

    def test_minimum_and_invalid_area_are_never_guessed(self):
        wb=load_workbook(self.path);wb['Step 2']['E15']=None;wb.save(self.path);wb.close()
        f=read_pc_facts(self.path);self.assertIsNone(f['qualifying_count']);self.assertTrue(f['warnings'])
        wb=load_workbook(self.path);wb['Step 2']['D2']=None;wb.save(self.path);wb.close()
        with self.assertRaisesRegex(ValueError,'positive'):read_pc_facts(self.path)

    def test_legacy_current_area_and_corner_flags(self):
        wb=load_workbook(self.path);s=wb['Step 2'];s['C2']='';s['D2']=None;s['E10']=5000
        for r in range(15,20):s.cell(r,4,s.cell(r,5).value);s.cell(r,5).value=None
        wb.save(self.path);wb.close()
        f=read_pc_facts(self.path);self.assertEqual(f['qualifying_count'],3);self.assertEqual(f['eligible_area'],16000)

    def test_report_uses_renewal_source_and_adoption_not_effective_date(self):
        ctx,notes=report_context(self.facts,{'condition_statement':'The block remains substantially unchanged.','notice_sent':True,'hearing_notice_sent':True})
        doc=fill_template(ctx,'mls_renewal');text='\n'.join(p.text for p in doc.paragraphs)
        self.assertIn('2006-121',text);self.assertIn('March 1, 2006',text);self.assertNotIn('March 15, 2006',text)
        self.assertIn('75% by lot count and 94% by land area',text)
        self.assertIn('Evidence of support is not required',text)
        self.assertNotIn('Cleburne',text);self.assertNotIn('2005-1287',text);self.assertNotIn('5,775',text)
        self.assertTrue(any('SharePoint source' in n for n in notes))
        ctx,notes=report_context(self.facts,{})
        self.assertIn('[Confirm',ctx['notice_sentence']);self.assertIn('[Enter',ctx['condition_statement'])

    def test_speech_does_not_invent_renewal_support(self):
        opts=defaults(self.facts,{})
        opts.update(presenter='Sample Planner',agenda='12',neighborhood='Example',district='C',super_neighborhood='Example',subdivision='Example',plat_year='1926',earliest_year='1930',protest_count='1',notice_mailed=True,recommendation='Approval')
        ctx,notes=speech_context(self.facts,opts);text='\n'.join(p.text for p in fill_speech(ctx).paragraphs)
        self.assertIn('renewal of a 5,000',text);self.assertIn('555REN',text);self.assertIn('2006-121',text)
        self.assertIn('petition support is not required',text);self.assertNotIn('owners of 0',text);self.assertNotIn('{{',text)
        self.assertFalse(notes)
        out=generate_speech(self.path,self.facts,opts,self.folder,export_pdf=False)
        self.assertTrue(Path(out['docx']).is_file())

    def test_standard_map_palette_and_renewal_identity(self):
        from tools.test_boundary_map import workbook, fixture
        from core.boundary_map import read_map_case
        # The standard fixture provides only geometry; renewal facts determine the PC markings.
        p=self.folder/'geometry.xlsx';workbook(p);data=fixture(read_map_case(p,'mls'))
        facts=copy.deepcopy(self.facts)
        for row,match in zip(facts['rows'],data['matched']):
            prop=match['property'];row['description']=f"{prop.get('number') or '0'} ({prop.get('legal_description','')})"
        data['case']['wizard']='mls_renewal';data['case']['reference']='MLS 555REN'
        data=attach_pc_rows(data,facts,[facts['rows'][1]['row']])
        self.assertFalse(any(m['pc']['signed'] for m in data['matched']))
        self.assertTrue(data['matched'][1]['pc']['protested'])
        rendered=render_pc_map(data,pc_defaults(data));text=' '.join(PdfReader(io.BytesIO(rendered['pdf'])).pages[0].extract_text().split())
        self.assertIn('Lot Size Block Renewal',text);self.assertIn('Property owner protesting application',text)
        self.assertEqual(len(rendered['legend_entries']),14)

    def test_report_saves_without_optional_attachments(self):
        opts={'include_calculation':False,'condition_statement':'Sample condition','notice_sent':True,'hearing_notice_sent':True}
        with patch('core.office_convert.office_to_pdf',side_effect=RuntimeError('Test without Office')):
            out=generate_pc_report(self.path,self.facts,opts,self.folder)
        text='\n'.join(p.text for p in Document(out['docx']).paragraphs)
        self.assertIn('555REN',text);self.assertIn('No attachments selected.',text);self.assertNotIn('{{',text)

    def test_renewal_work_area_and_speech_open_from_files_based_owner(self):
        import tkinter as tk
        import time
        from wizards.mls.pc_ui import PCWorkspace
        from wizards.mls.pc_speech_ui import PCSpeech
        root=tk.Tk();root.withdraw();self.addCleanup(root.destroy)
        root.files={'workbook':str(self.path),'out':str(self.folder),'map':''}
        root._boundary_map_wizard_key='mls_renewal'
        errors=[]
        with patch('core.workbooks.WorkbookReadCache.resolve',return_value=str(self.path)), patch('tkinter.messagebox.showerror',side_effect=lambda *a,**k:errors.append(a)):
            workspace=PCWorkspace(root);speech=PCSpeech(root)
            limit=time.time()+12
            while time.time()<limit and (not workspace.facts or not speech.facts):root.update();time.sleep(.05)
            self.assertFalse(errors)
            self.assertEqual(workspace.facts['application'],'555REN')
            self.assertEqual(speech.facts['application'],'555REN')
            self.assertEqual(workspace.fields['ordinance'].get(),'2006-121')
            self.assertEqual(speech.fields['ordinance'].get(),'2006-121')
            self.assertIn('renewal of a 5,000',speech.preview.get('1.0','end'))
            self.assertEqual(workspace._boundary_map_wizard_key,'mls_renewal')
            workspace.close();speech.close()


if __name__=='__main__':unittest.main()
