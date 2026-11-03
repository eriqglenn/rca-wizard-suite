"""Independent review fixtures: actual documents with deliberate discrepancies."""
import copy
import hashlib
import importlib
import io
import json
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import patch
from openpyxl import Workbook
from docx import Document
import fitz
from core import peer_review as pr
from core import peer_readers as readers


def workbook(path,key='mls'):
    book=Workbook();s=book.active;s.title='Step 1 - Input Properties Info'
    s.append(['TAX_ID','Site_addr_1','Legal_Dscr_1','Legal_Dscr_2','Total_Land_Area','Land Use','OWNER_MAILTO','Mail_Addr_1','Property Description'])
    s.append(['0350820010001','100 TEST ST','LT 1 BLK 1','TEST',5000,'SFR','TEST OWNER','100 MAIL ST','100 (LT 1 BLK 1)'])
    s.append(['0350820010001','100 TEST ST','LT 2 BLK 1','TEST',5000,'SFR','TEST OWNER','100 MAIL ST','100 (LT 2 BLK 1)'])
    s.append(['0350820010002','102 TEST ST','LT 3 BLK 1','TEST',10000,'VAC','OTHER OWNER','102 MAIL ST','102 (LT 3 BLK 1)'])
    s2=book.create_sheet('Step 2 - Calculation Sheet');renewal=key.endswith('_renewal');family=key.split('_')[0]
    s2['A1']='SPECIAL MINIMUM '+('LOT SIZE' if family=='mls' else 'BUILDING LINE')+' BLOCK'+(' RENEWAL' if renewal else '')
    for c,v in {'B2':'187REN' if renewal else '915','B4':'Test Street','D4':'100 block of Test Street','B5':'First Street','D5':'Second Street','B6':'North'}.items():s2[c]=v
    minimum=5000 if family=='mls' else 20
    s2['D2' if renewal else 'L16' if family=='mbl' else 'K16']=minimum
    s2['D9' if renewal else 'H11']=.67;s2['A13' if renewal else 'A15']=3
    if not renewal:s2['F8']=.25
    header=14 if renewal else 16
    for col,val in enumerate(['Property Description','Land Use','Original Building Line' if renewal and family=='mbl' else 'Signed in Support','Current Building Line' if family=='mbl' and renewal else 'Original Lot Size' if renewal else 'Lot size (in Sq Feet)','Lot size' if renewal else 'Building Line','Meets'],1):s2.cell(header,col,val)
    for i,(desc,lu,area) in enumerate([('100 (LT 1 BLK 1)','SFR',5000),('100 (LT 2 BLK 1)','SFR',5000),('102 (LT 3 BLK 1)','VAC',10000)],header+1):
        if renewal:values=[desc,lu,20 if family=='mbl' else '',20 if family=='mbl' else area,area,'YES']
        else:values=[desc,lu,'Y' if i==header+1 else '',area,20 if family=='mbl' and lu=='SFR' else 0]
        for c,v in enumerate(values,1):s2.cell(i,c,v)
    if not renewal:s2['L17' if family=='mbl' else 'K17']='=IF(H17>=0.7,F17,"")'
    ci=book.create_sheet('Case Info');ci.append(['Original_Ordinance_No','2006-1019']);ci.append(['Original_Effective_Date','10/10/2006']);ci.append(['Original_Expiration_Date','10/10/2026']);ci.append(['Council_District','H'])
    book.save(path);book.close();return path


def pdf(path,text):
    with fitz.open() as doc:
        p=doc.new_page();p.insert_textbox(fitz.Rect(35,35,570,770),text,fontsize=10);doc.save(path)
    return path


def hcad(account='0350820010001',area=10000):
    return f'HCAD Property Search\nAccount Number: {account}\nSite Address: 100 TEST ST\nName: TEST OWNER\nMailing Address: 100 MAIL ST\nLegal Description: LTS 1 & 2 BLK 1\nLand Area: {area:,} SF\nLiving Area: 2200 SF'


class PeerReviewTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup);self.root=Path(self.tmp.name)

    def test_all_four_calculation_contexts(self):
        for key in ('mls','mbl','mls_renewal','mbl_renewal'):
            with self.subTest(key=key):
                facts=readers.read_calculation(workbook(self.root/(key+'.xlsx'),key));r=pr.Review(key);pr.calculation_checks(r,facts)
                self.assertEqual(key,facts['wizard']);self.assertEqual(3,facts['metrics']['eligible_count'])
                self.assertEqual(2,facts['metrics']['sfr_count'])
                self.assertFalse([i for i in r.result['findings'] if i['severity']=='Error'],r.result['findings'])
                if not facts['renewal']:self.assertEqual(25,facts['metrics']['support_area_percent'])

    def facts(self,key='mls'):
        f=readers.read_calculation(workbook(self.root/(key+'.xlsx'),key));r=pr.Review(key);pr.calculation_checks(r,f);return r,f

    def block(self,text,method='PDF text'):
        return {'text':text,'location':'test.pdf, page 2','method':method}

    def test_hcad_split_account_sum_and_duplicate_identity(self):
        r,f=self.facts();pr.hcad_checks(r,f,[self.block(hcad())],[])
        sums=[x for x in r.result['findings'] if x['title']=='HCAD land area / combined lot rows']
        self.assertEqual(['Passed'],[x['severity'] for x in sums])
        self.assertFalse(any(x['title']=='Duplicate account and legal lot row' for x in r.result['findings']))
        r=pr.Review('mls');pr.hcad_checks(r,f,[self.block(hcad(area=15000))],[])
        self.assertTrue(any(x['title']=='HCAD land area / combined lot rows' and x['severity']=='Error' for x in r.result['findings']))

    def test_hcad_current_viewer_columns_not_building_area(self):
        text='Account: 0350820010001\nLegal\nDescription\nLand\nLiving Area\nFiduciary\nLTS 1 & 2 BLK 1 TEST\n10,000 SF\n2,500 SF\nLocation'
        record=readers.hcad_record(text);self.assertEqual(10000,record['area']);self.assertEqual('LTS 1 & 2 BLK 1 TEST',record['legal'])
        self.assertIsNone(readers.hcad_record('Account: 0350820010001\nLiving Area: 2500 SF')['area'])
        expanded=text.replace('Living Area\nFiduciary','Building Area\nBuilding Class\nUnits\nNet Rentable\nArea\nFiduciary\nStatus')
        self.assertEqual(10000,readers.hcad_record(expanded)['area'])

    def test_conflicting_hcad_years_do_not_choose_area(self):
        r,f=self.facts();pr.hcad_checks(r,f,[self.block(hcad()),self.block(hcad(area=15000))],[])
        self.assertTrue(any(x['title']=='Conflicting HCAD versions' for x in r.result['findings']))
        self.assertFalse(any(x['title']=='HCAD land area / combined lot rows' for x in r.result['findings']))

    def test_ordinance_count_words_minimum_and_legal_blanks(self):
        r,f=self.facts()
        text='fifteen (14) out of three (3) of the lots within the Application Area are at least five thousand (5,000) square feet. The minimum lot size is 6,000 square feet. PASSED AND ADOPTED ON ______. EXHIBIT A EXHIBIT B EXHIBIT C'
        pr.document_checks(r,f,[self.block(text)],'Draft ordinance')
        errors={x['title'] for x in r.result['findings'] if x['severity']=='Error'}
        self.assertIn('Written number / numeral',errors);self.assertIn('Lots meeting proposed minimum',errors);self.assertIn('Proposed minimum in Draft ordinance',errors)
        self.assertNotIn('Unresolved document placeholder',errors)

    def test_mbl_minimum_is_not_lot_area(self):
        r,f=self.facts('mbl_renewal')
        pr.document_checks(r,f,[self.block('The lot size is 10,000 square feet. The minimum building line of twenty (20) feet shall be established.')],'Draft ordinance')
        items=[i for i in r.result['findings'] if i['title']=='Proposed minimum in Draft ordinance']
        self.assertEqual([20],[i['actual'] for i in items])

    def test_ocr_findings_are_never_confirmed_passes_or_errors(self):
        r,f=self.facts();r.result['findings']=[]
        pr.document_checks(r,f,[self.block('minimum lot size of 6,000 square feet. X (X)','OCR')],'Draft ordinance')
        self.assertFalse(any(x['severity'] in ('Error','Passed') for x in r.result['findings']))

    def test_percentage_denominators_and_thresholds(self):
        r,f=self.facts();r.result['findings']=[]
        for text in ['SFR: 67%.','25% of the eligible application area signed in support.','SFR share must be at least 60%.']:
            pr.document_checks(r,f,[self.block(text)],'RCA')
        self.assertFalse(any(x['severity']=='Error' for x in r.result['findings']),r.result['findings'])
        pr.document_checks(r,f,[self.block('SFR: 25%.')],'RCA')
        self.assertTrue(any(x['title']=='Sfr lot percent' and x['severity']=='Error' for x in r.result['findings']))

    def test_renewal_dates_and_original_ordinance(self):
        r,f=self.facts('mls_renewal');pr.document_checks(r,f,[self.block('Ordinance No. 2006-999 effective on October 11, 2006; Ordinance No. 2006-999 expired on October 10, 2026.')],'Draft ordinance')
        errors={x['title'] for x in r.result['findings'] if x['severity']=='Error'}
        self.assertIn('Previous ordinance reference',errors);self.assertIn('Original Effective Date',errors)

    def test_pdf_docx_tables_and_placeholder_evidence(self):
        doc=Document();doc.add_paragraph('Case MLS 915');doc.add_table(rows=1,cols=1).cell(0,0).text='{{missing}}'
        path=self.root/'draft.docx';doc.save(path);blocks=readers.read_document(path,use_ocr=False)
        self.assertTrue(any('{{missing}}' in b['text'] for b in blocks));self.assertTrue(all('paragraph' in b['location'] for b in blocks))
        path=pdf(self.root/'rca.pdf','minimum lot size of 5,000 square feet. Test Street, First Street and Second Street.')
        blocks=readers.read_document(path,use_ocr=False);self.assertEqual('page 1',blocks[0]['location'])

    def test_end_to_end_sources_unchanged_and_rerun(self):
        inputs={'calculation':str(workbook(self.root/'calc.xlsx')),'rca':str(pdf(self.root/'rca.pdf','Special Minimum Lot Size\n100 block of Test Street\nFirst Street and Second Street\n5,000 Square Feet\nSource: Harris Central Appraisal District')),'hcad':[str(pdf(self.root/'hcad.pdf',hcad()))]}
        doc=Document();doc.add_paragraph('MLS 915. The minimum lot size is 6,000 square feet.');doc.save(self.root/'draft.docx');inputs['ordinance']=str(self.root/'draft.docx')
        before={p:hashlib.sha256(Path(p).read_bytes()).hexdigest() for p in [inputs['calculation'],inputs['rca'],inputs['hcad'][0],inputs['ordinance']]}
        result=pr.run_review(inputs,'mls',use_ocr=False)
        self.assertTrue(any(x['title']=='Proposed minimum in Draft ordinance' and x['severity']=='Error' for x in result['findings']))
        self.assertTrue(any(x['title']=='Boundary and address placement' and x['severity']=='Needs review' for x in result['findings']))
        self.assertEqual(before,{p:hashlib.sha256(Path(p).read_bytes()).hexdigest() for p in before})
        doc.paragraphs[0].text='MLS 915. The minimum lot size is 5,000 square feet.';doc.save(self.root/'draft.docx')
        new=pr.run_review(inputs,'mls',use_ocr=False)
        self.assertFalse(any(x['title']=='Proposed minimum in Draft ordinance' and x['severity']=='Error' for x in new['findings']))
        report=self.root/'review.html';pr.save_report(new,report);self.assertIn('page 1',report.read_text(encoding='utf8'))
        with self.assertRaises(ValueError):pr.save_report(new,inputs['calculation'])

    def test_missing_invalid_cancel_and_html_escaping(self):
        result=pr.run_review({'calculation':str(self.root/'missing.xlsx')},'mls',use_ocr=False)
        self.assertEqual(0,result['counts'].get('Passed',0))
        result['findings'].append({'severity':'Error','title':'<script>alert(1)</script>','details':'bad & data','locations':['cell <A1>']})
        self.assertNotIn('<script>',pr.report_html(result))
        cancel=threading.Event();cancel.set()
        with self.assertRaises(InterruptedError):pr.run_review({},'mls',cancel=cancel)

    def test_sorted_rows_and_excluded_zero_area(self):
        r,f=self.facts();f['table'].reverse();f['table'][0]['land_use']='EXC';f['table'][0]['area']=0
        f['cached']['A15']=2;f['cached']['H11']=1;f['cached']['F8']=.5
        r=pr.Review('mls');pr.calculation_checks(r,f)
        self.assertFalse(any(i['severity']=='Error' for i in r.result['findings']),r.result['findings'])
        self.assertTrue(any(i['title']=='Excluded lot retained in review' for i in r.result['findings']))

    def test_missing_measurement_and_unknown_land_use_do_not_pass_metrics(self):
        r,f=self.facts('mbl');f['table'][0]['measure']=None;r=pr.Review('mbl');pr.calculation_checks(r,f)
        self.assertNotIn('qualifying_count',f['metrics'])
        f['table'][0]['land_use']='UNKNOWN';pr.calculation_checks(pr.Review('mbl'),f)
        self.assertNotIn('sfr_lot_percent',f['metrics'])

    def test_formula_errors_and_uncached_minimum_are_reported(self):
        from openpyxl import load_workbook
        path=workbook(self.root/'broken.xlsx');b=load_workbook(path);s=b['Step 2 - Calculation Sheet'];s['K16']='=1/0';s['A15']='#NUM!';b.save(path);b.close()
        f=readers.read_calculation(path);r=pr.Review('mls');pr.calculation_checks(r,f)
        self.assertTrue(any(i['title']=='Spreadsheet formula error' for i in r.result['findings']))
        self.assertTrue(any(i['title']=='Proposed minimum is unavailable' for i in r.result['findings']))

    def test_detect_changed_source_after_snapshot(self):
        path=workbook(self.root/'changing.xlsx');original=readers.read_calculation
        def change(snapshot):
            facts=original(snapshot);path.write_bytes(path.read_bytes()+b'changed after snapshot');return facts
        with patch.object(pr,'read_calculation',side_effect=change):result=pr.run_review({'calculation':str(path)},'mls',use_ocr=False)
        self.assertEqual('stale',result['status'])

    def test_renewal_unchanged_lot_not_counted_as_meeting_minimum(self):
        r,f=self.facts('mls_renewal');f['table'][0].update(measure=4000,original=4000,flag='YES')
        pr.calculation_checks(pr.Review('mls_renewal'),f)
        self.assertEqual(2,f['metrics']['qualifying_count']);self.assertEqual(3,f['metrics']['unchanged_or_meets_count'])

    def test_packet_ocr_account_coverage_not_marked_passed(self):
        r,f=self.facts();pr.hcad_checks(r,f,[],[self.block(hcad(),'OCR'),self.block(hcad('0350820010002'),'OCR')])
        coverage=[i for i in r.result['findings'] if i['title']=='HCAD account coverage in RCA']
        self.assertEqual('Needs review',coverage[0]['severity'])

    def test_all_four_sidebar_buttons_and_actual_review_controls(self):
        import tkinter as tk
        from core import session
        root=tk.Tk();root.withdraw();self.addCleanup(root.destroy)
        classes={'mls':'MailoutWizard','mbl':'MBLWizard','mls_renewal':'MLSRenewalWizard','mbl_renewal':'MBLRenewalWizard'}
        with patch.object(session,'load',return_value={}),patch.object(session,'save'):
            for key,name in classes.items():
                with self.subTest(key=key):
                    owner=getattr(importlib.import_module('wizards.'+key+'.wizard'),name)(root)
                    self.assertEqual(5,len(owner.step_buttons))
                    self.assertIn('Generate',owner.step_buttons[3]['text'])
                    self.assertEqual('5. Peer Reviewer',owner.step_buttons[4]['text'])
                    owner.step_buttons[4].invoke();pane=owner._peer_reviewer
                    root.update();pane.vars['calculation'].set(str(workbook(self.root/(key+'_ui.xlsx'),key)));pane.ocr.set(False)
                    for k in ('rca','hcad','ordinance'):pane.vars[k].set('')
                    pane.run.invoke();self.assertTrue(pane.running)
                    deadline=time.time()+20
                    while pane.running and time.time()<deadline:root.update();time.sleep(.01)
                    self.assertFalse(pane.running);self.assertIsNotNone(pane.result);self.assertEqual('normal',str(pane.save['state']))
                    self.assertTrue(pane.tree.get_children())
                    pane.vars['rca'].set('new.pdf');self.assertIsNone(pane.result);self.assertEqual('disabled',str(pane.save['state']))
                    owner.destroy();root.update()


def main():
    return 0 if unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(PeerReviewTests)).wasSuccessful() else 1


if __name__=='__main__':
    raise SystemExit(main())
