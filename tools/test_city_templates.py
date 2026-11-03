"""City source preservation and context-sensitive field regression tests."""
import importlib,tempfile,unittest
from pathlib import Path
from docx import Document
from openpyxl import Workbook
from core.city_templates import field_values,render_updated
from core.paths import templates_dir

class CityTemplateTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.folder=Path(self.temp.name)
        self.wb=self.folder/'case.xlsx';w=Workbook();s=w.active;s.title='Case Info'
        for row in [('Original_Passage_Date','2005-12-14'),('Original_Effective_Date','2006-01-31'),('PC_Noncorner_Lot_Count',2),('Subdivision_Plat_Year',1927),('Structure_Age_Statement','Homes date to the 1950s.'),('Council_Deadline','2027-01-12')]:s.append(row)
        w.save(self.wb)
        self.ctx=dict(Application_Number='712',Application_Renewal_Number='712REN',Minimum_Lot_Size_Text='5,000',Minimum_Lot_Size_Words='five thousand',Minimum_Building_Line_Text='22',Minimum_Building_Line_Words='twenty-two',Area_Description='1700 block of Example Street, both sides, between First and Second Streets',Area_Line_1='1700 block of Example Street, both sides,',Area_Line_2='between First and Second Streets',Council_District='C',Council_Member='Example Member',Planner_Name='Test Planner',Planner_Phone='832-393-6637',Original_Ordinance_No='2005-999',Approval_Path='Director Approval',Protest_Count=0,Total_Lot_Count=3,Qualifying_Lot_Count=2,Developed_Lot_Count=2,Support_Percentage='25',SFR_Percentage='67',Property_Rows=[dict(desc='LT 1',land_use='SFR',lot=5000,bl=22,cur=5000,signed=True,account='same'),dict(desc='LT 2',land_use='SFR',lot=5000,bl=23,cur=5000,signed=False,account='same'),dict(desc='LT 3',land_use='VAC',lot=10000,bl=0,cur=10000,signed=False,account='other'),dict(desc='EXC',land_use='EXC',lot=90000,bl=99,cur=90000,signed=True)])
    def tearDown(self):self.temp.cleanup()
    def read(self,p):return '\n'.join(''.join(p.xpath('.//w:t/text()')) for p in Document(p).element.xpath('.//w:p'))
    def test_area_and_lot_denominators_split_lots_and_exclusions(self):
        for wizard in ('mls','mbl','mls_renewal'):
            f=field_values(self.ctx,self.wb,wizard)
            self.assertEqual(f['TOTAL'],'3');self.assertEqual(f['SFR_LOT_PERCENT'],'67')
            self.assertEqual(f['SFR_AREA_PERCENT'],'50');self.assertEqual(f['SUPPORT_AREA_PERCENT'],'25')
            self.assertEqual(f['QUALIFYING'],'2' if wizard=='mbl' else '3')
            self.assertEqual(f['PASSAGE_YEAR'],'2005');self.assertEqual(f['EFFECTIVE_DATE'],'January 31, 2006')
    def test_unknown_is_not_replaced_with_effective_date(self):
        from openpyxl import load_workbook
        w=load_workbook(self.wb);w.active['B1']=None;w.save(self.wb)
        self.assertIn('Enter Original_Passage_Date',field_values(self.ctx,self.wb,'mbl_renewal')['PASSAGE_DATE'])
    def test_source_units_and_case_fields(self):
        cases=[('mls','PC_Approval_NEW.docx','pc'),('mbl','MBL_PC_Approval.docx','pc'),('mls_renewal','MLS_REN_PC_Approval.docx','pc'),('mls_renewal','MLS_REN_Director_Approval.docx','director')]
        for wizard,file,role in cases:
            out=self.folder/file
            self.assertTrue(render_updated(templates_dir(wizard)/file,out,self.ctx,self.wb,wizard,role))
            t=self.read(out);self.assertNotIn('{{',t);self.assertNotIn('between X Street',t)
            if wizard=='mls_renewal' and role=='pc':
                self.assertIn('two (2) lots (excluding corner lots)',t);self.assertIn('5,000 square feet',t);self.assertNotIn('Building Line',t)
            if wizard=='mbl':self.assertIn('22 ft',t);self.assertIn('50% of the proposed application area',t)
            if wizard=='mls':self.assertIn('1927',t);self.assertIn('Homes date to the 1950s.',t)
    def test_routing_checkbox_paths_and_source_metrics(self):
        for wizard,file in [('mls','MLS_Routing Form.docx'),('mbl','MBL_Routing_Form.docx'),('mls_renewal','MLS_Renewal_Routing_Form.docx'),('mbl_renewal','MBL_REN_Routing_Form.docx')]:
            for pc in (False,True):
                ctx=dict(self.ctx,Approval_Path='Planning Commission Approval' if pc else 'Director Approval')
                out=self.folder/(wizard+str(pc)+'.docx');render_updated(templates_dir(wizard)/file,out,ctx,self.wb,wizard,'routing');t=self.read(out)
                self.assertNotIn('{{',t)
                if 'Contains' in t:self.assertIn('Yes ☒ No ☐',t)
                self.assertIn(('☒' if pc else '☐')+' Planning Commission Approval',t)
                self.assertIn(('☒' if wizard.endswith('renewal') else '☐')+' Original Ordinance',t)
                if wizard=='mbl':self.assertIn('2 out of 3 of the lots are at least 22 feet',t)
    def test_staff_uses_adoption_date_and_correct_variants(self):
        for wizard,file in [('mls_renewal','MLS_REN_OneMeeting_Director_Approval.docx'),('mbl_renewal','MBL_REN_OneMeeting_Staff_Report.docx')]:
            out=self.folder/(wizard+'.docx');render_updated(templates_dir(wizard)/file,out,self.ctx,self.wb,wizard,'staff');t=self.read(out)
            self.assertIn('December 14, 2005',t);self.assertNotIn('January 31, 2006',t)
        for w in ('mls','mbl'):
            out=self.folder/(w+'.docx');render_updated(templates_dir(w)/(w.upper()+'_OneMeeting_Staff_Report.docx'),out,dict(self.ctx,Protest_Count=2),self.wb,w,'staff')
            self.assertIn('Two (2) protests were received.',self.read(out))
    def test_legacy_custom_templates_keep_existing_filler(self):
        p=self.folder/'custom.docx';Document().save(p)
        self.assertFalse(render_updated(p,self.folder/'out.docx',self.ctx,self.wb,'mls','pc'))
