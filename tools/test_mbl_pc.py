"""MBL and renewal PC integration and measure regression tests."""
import copy,io,tempfile,unittest,time
from pathlib import Path
from unittest.mock import patch
from openpyxl import Workbook,load_workbook
from docx import Document
from pypdf import PdfReader
from wizards.mls.pc_data import read_pc_facts,attach_pc_rows
from wizards.mbl.pc_report import report_context
from wizards.mbl.pc_speech import speech_context
from wizards.mls.pc_report import fill_template,generate_pc_report
from wizards.mls.pc_speech import fill_speech,generate_speech
from wizards.mls.pc_map import render_pc_map,pc_defaults

def case_fixture(folder,renewal=False):
    p=Path(folder)/('mbl_renewal.xlsx' if renewal else 'mbl.xlsx')
    wb=Workbook();s1=wb.active;s1.title='Step 1';s1.append(['Property Description','group_dscr'])
    s=wb.create_sheet('Step 2');s['A1']='SPECIAL MINIMUM BUILDING LINE BLOCK'+(' RENEWAL' if renewal else '')
    for c,v in dict(B2='121',B4='East 24th Street',D4='300 block of East 24th Street',B5='Cortlandt Street',D5='Arlington Street',B6='north',D6='south').items():s[c]=v
    header=14 if renewal else 16
    if renewal:s['C2']='Building Line to Renew (in Feet)';s['D2']=15
    else:s['E4']='15-Foot';s['L16']=15
    for c,v in enumerate(['Property Description','Land Use','Original Building Line','Current Building Line','Lot Size','Meets Minimum'] if renewal else ['Property Description','Land Use','Signed in Support','Lot Size','Building Line'],1):s.cell(header,c,v)
    for i,(lu,area,bl,signed) in enumerate([('SFR',5000,20,'Y'),('SFR',5000,10,'Y'),('SFR',5000,15,'N'),('COM',1000,0,'N'),('VAC',4000,0,'Y'),('EXC',9000,5,'Y')],header+1):
        desc=f'{300+i} (LT {i-header} BLK 1)'
        values=[desc,lu,30,bl,area,'YES'] if renewal else [desc,lu,signed,area,bl]
        for c,v in enumerate(values,1):s.cell(i,c,v)
        s1.append([desc,'Single-Family Residential' if lu=='SFR' else lu])
    ci=wb.create_sheet('Case Info')
    for pair in [('Original_Ordinance_No','2006-121'),('Original_Passage_Date','March 1, 2006'),('Original_Effective_Date','March 15, 2006'),('PC_Agenda_Number','12'),('PC_Hearing_Date','October 8, 2026'),('Protest_Count','1'),('Subdivision_Name','Example')]:ci.append(pair)
    wb.save(p);wb.close();return p

def options():
    return dict(agenda='12',hearing_date='October 8, 2026',protest_count='1',plat_year='1926',structure_age='Most houses date from the 1930s.',preservation='The setbacks retain the original development pattern.',notice_sent=True,hearing_notice_sent=True,notice_mailed=True,presenter='Sample Planner',neighborhood='Example',district='C',super_neighborhood='Example',subdivision='Example',earliest_year='1930',recommendation='Approval',include_calculation=False)

class MBLPCTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup);self.folder=Path(self.temp.name)
        self.paths={k:case_fixture(self.folder,k.endswith('renewal')) for k in ('mbl','mbl_renewal')}

    def facts(self,key='mbl'):return read_pc_facts(self.paths[key],wizard_key=key)

    def test_area_vs_counts_current_vs_original_and_split_rows(self):
        for k in self.paths:
            f=self.facts(k)
            self.assertEqual((f['eligible_count'],f['excluded_count'],f['developed_count'],f['vacant_count']),(5,1,4,1))
            self.assertEqual((f['sfr_lot_percent'],f['sfr_area_percent'],f['qualifying_count'],f['developed_percent']),('60','75',2,'50'))
            self.assertEqual(f['minimum'],15)
        self.assertEqual(self.facts()['support_area_percent'],'70')
        f=self.facts('mbl_renewal');self.assertIsNone(f['signed_count']);self.assertFalse(any(r['signed'] for r in f['rows']))

    def test_missing_measurement_and_minimum_not_guessed(self):
        p=self.paths['mbl'];wb=load_workbook(p);wb['Step 2']['E17']=None;wb.save(p);wb.close()
        self.assertIsNone(self.facts()['qualifying_count'])
        wb=load_workbook(p);wb['Step 2']['L16']=None;wb['Step 2']['E4']=None;wb.save(p);wb.close()
        with self.assertRaisesRegex(ValueError,'positive'):self.facts()

    def test_legacy_renewal_and_fractional_feet(self):
        p=self.paths['mbl_renewal'];wb=load_workbook(p);s=wb['Step 2'];s['C2']='';s['D2']=None;s['E10']=15.5;wb.save(p);wb.close()
        f=self.facts('mbl_renewal');self.assertEqual(f['minimum_text'],'15.5');self.assertEqual(f['qualifying_count'],1)

    def test_invalid_area_blocks_percentages_not_building_line_counts(self):
        p=self.paths['mbl'];wb=load_workbook(p);wb['Step 2']['D17']=-1;wb.save(p);wb.close()
        f=self.facts();self.assertEqual(f['sfr_area_percent'],'[review]');self.assertEqual(f['qualifying_count'],2)

    def test_wrong_wizard_rejected(self):
        with self.assertRaises(ValueError):read_pc_facts(self.paths['mbl'],wizard_key='mbl_renewal')
        with self.assertRaises(ValueError):read_pc_facts(self.paths['mbl_renewal'],wizard_key='mbl')

    def test_report_speech_source_words_units_and_dates(self):
        for key in self.paths:
            f=self.facts(key);ctx,notes=report_context(f,options());ctx['attachment_list']='No attachments selected.'
            text='\n'.join(p.text for p in fill_template(ctx,key).paragraphs)
            self.assertIn('15 feet',text);self.assertIn('75% of the eligible application land area',text)
            self.assertIn('forty years',text);self.assertNotIn('{{',text);self.assertNotIn('1917',text)
            if key.endswith('renewal'):
                self.assertIn('March 1, 2006',text);self.assertNotIn('March 15, 2006',text)
                self.assertIn('Petition support is not recorded',text);self.assertIn('source template',text)
            ctx,notes=speech_context(f,options());text='\n'.join(p.text for p in fill_speech(ctx).paragraphs)
            self.assertIn('15-foot Special Minimum Building Line Block',text);self.assertIn('2 of 4 developed lots (50%)',text)
            self.assertNotIn('square feet',text);self.assertNotIn('Lot Size',text);self.assertNotIn('{{',text)

    def test_standard_map_units_flags_and_palette(self):
        from tools.test_boundary_map import workbook,fixture
        from core.boundary_map import read_map_case
        for key in self.paths:
            p=self.folder/'geometry.xlsx';workbook(p);data=fixture(read_map_case(p,'mls'));f=self.facts(key)
            for row,m in zip(f['rows'],data['matched']):
                prop=m['property'];row['description']=f"{prop.get('number') or '0'} ({prop.get('legal_description','')})"
            data['case'].update(wizard=key,reference='MBL '+f['application'],minimum='15 Feet')
            data=attach_pc_rows(data,f,[f['rows'][1]['row']]);self.assertTrue(data['matched'][1]['pc']['below']);self.assertFalse(data['matched'][0]['pc']['below'])
            r=render_pc_map(data,pc_defaults(data));text=' '.join(PdfReader(io.BytesIO(r['pdf'])).pages[0].extract_text().split())
            self.assertIn('Building Line Block',text);self.assertIn('Minimum Building Line: 15 Feet',text);self.assertNotIn('Minimum Lot Size',text)
            self.assertEqual(len(r['legend_entries']),14)
            if key.endswith('renewal'):self.assertIn('Renewal',text);self.assertFalse(any(m['pc']['signed'] for m in data['matched']))

    def test_save_without_optional_attachments(self):
        for key,p in self.paths.items():
            f=self.facts(key)
            with patch('core.office_convert.office_to_pdf',side_effect=RuntimeError('Office unavailable in this test')):
                out=generate_pc_report(p,f,options(),self.folder)
            self.assertTrue(Path(out['docx']).is_file());self.assertIn('MBL',Path(out['docx']).name)
            out=generate_speech(p,f,options(),self.folder,export_pdf=False);self.assertTrue(Path(out['docx']).is_file())

    def test_ui_both_wizards(self):
        import tkinter as tk
        from wizards.mls.pc_ui import PCWorkspace
        from wizards.mls.pc_speech_ui import PCSpeech
        root=tk.Tk();root.withdraw();self.addCleanup(root.destroy)
        for key,p in self.paths.items():
            root.files={'workbook':str(p),'out':str(self.folder),'map':''};root._boundary_map_wizard_key=key;errors=[]
            with patch('core.workbooks.WorkbookReadCache.resolve',return_value=str(p)),patch('tkinter.messagebox.showerror',side_effect=lambda *a,**kw:errors.append(a)):
                work=PCWorkspace(root);speech=PCSpeech(root);limit=time.time()+12
                while time.time()<limit and (not work.facts or not speech.facts):root.update();time.sleep(.04)
                self.assertFalse(errors);self.assertEqual(work.facts['wizard_key'],key);self.assertEqual(speech.facts['wizard_key'],key)
                self.assertIn('feet',work.metrics.get('1.0','end'));self.assertIn('15-foot',speech.preview.get('1.0','end'))
                self.assertIn('building_line',work.tree['columns']);work.close();speech.close()

if __name__=='__main__':unittest.main()
