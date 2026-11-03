"""Lot count/area, workbook preservation, geometry, roster and notification regressions."""
import copy,io,json,hashlib,tempfile,unittest,zipfile
from pathlib import Path
from unittest.mock import patch
from openpyxl import load_workbook
from shapely.geometry import box,mapping
from core.lot_divisions import *
from core.application_import import save_calculation
from core.workbooks import calculation_template_path
from core.council_district import parse_members,add_members,write_case_district_copy
from tools.test_application_import import fields,parcel

class LotsTests(unittest.TestCase):
    def setUp(self):self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup);self.folder=Path(self.tmp.name)
    def parent(self):
        p=parcel(area=6000);p['legal_description']='LTS 1 & 2 BLK 69';return p
    def split_parent(self):
        p=self.parent();p['lots']=proposed_lots(p)
        for l in p['lots']:l.update(area=3000,area_source='Verified test plat')
        return p
    def save(self,p,key='mls'):
        f={**fields(),'renewal_minimum':'3000'};f['application_number']='910REN' if key.endswith('renewal') else '910'
        return save_calculation(self.folder/(key+'.xlsx'),f,{'parcels':[p],'warnings':[],'sources':[]},{},key,reviewed=True)
    def test_lists_and_ranges(self):
        for s in ('LTS 1,2,3 BLK 4','LTS 1 & 2 & 3 BLK 4','LOTS 1-3 BLK 4','LTS 1 THRU 3 BLK 4'):
            self.assertEqual(['1','2','3'],[l['number'] for l in parse_lots(s)])
        for s in ('TRS 1 & 2 BLK 4','W 25 FT OF LTS 1 & 2 BLK 4','LTS 1 & 2 LESS E 5 FT BLK 4','LTS 3-1 BLK 4','LTS 1 & 1 BLK 4'):
            self.assertEqual([],parse_lots(s))
    def test_equal_areas_do_not_invent_boundary_geometry(self):
        lots=proposed_lots(self.parent());self.assertEqual([3000,3000],[l['area'] for l in lots]);self.assertFalse(any('geometry' in l for l in lots))
        self.assertEqual([3000,3000],[p['land_area'] for p in expand_parcels([self.parent()])])
    def test_unequal_areas_and_duplicate_lots(self):
        p=self.split_parent();p['lots'][0]['area']=2000;p['lots'][1]['area']=4000
        self.assertEqual([2000,4000],[r['land_area'] for r in expand_parcels([p])])
        p['lots'][1]['area']=3000
        with self.assertRaises(ValueError):expand_parcels([p])
        p=self.split_parent();p['lots'][1]['description']=p['lots'][0]['description']
        with self.assertRaises(ValueError):expand_parcels([p])
    def test_import_each_wizard_preserves_artwork_and_separates_lots(self):
        for key in ('mls','mbl','mls_renewal','mbl_renewal'):
            with self.subTest(key=key):
                path=self.save(self.split_parent(),key);w=load_workbook(path,read_only=True,data_only=False)
                s=w.worksheets[0];self.assertEqual([3000,3000],[s['AO2'].value,s['AO3'].value]);self.assertEqual(s['AX2'].value,s['AX3'].value)
                self.assertEqual(['LT 1 BLK 69','LT 2 BLK 69'],[s['AD2'].value,s['AD3'].value]);self.assertEqual('hidden',w[META].sheet_state)
                self.assertEqual(2,len(read_metadata(w)['rows']));self.assertEqual(len(s[1]),56)
                if key=='mls_renewal':self.assertIsNone(w.worksheets[1]['D15'].value);self.assertTrue(w.worksheets[1]['E15'].value.startswith('='))
                w.close()
                with zipfile.ZipFile(calculation_template_path(key)) as before,zipfile.ZipFile(path) as after:
                    for name in before.namelist():
                        if name.startswith(('xl/media/','xl/drawings/','xl/styles')):self.assertEqual(before.read(name),after.read(name),name)
    def test_existing_copy_updates_rows_preserves_rank_and_formulas(self):
        source=self.save(parcel());rows,meta=workbook_rows(source);rows[0].update(legal_description='LTS 1 & 2 BLK 69',land_area=6599)
        rows[0]['lots']=[{'description':'LT 1 BLK 69','area':3000,'land_use':'SFR'},{'description':'LT 2 BLK 69','area':3599,'land_use':'VAC'}]
        before=source.read_bytes();out=self.folder/'split.xlsx';reviewed_workbook_copy(source,out,rows,meta,'mls')
        self.assertEqual(before,source.read_bytes());w=load_workbook(out,read_only=True,data_only=False);s=w.worksheets[1]
        self.assertEqual([1,2,3],[s[f'E{i}'].value for i in (17,18,19)])
        self.assertEqual("='Step 1 - Input Properties Info'!BD3",s['D18'].value)
        self.assertEqual("='Step 1 - Input Properties Info'!BB3",s['A18'].value)
        self.assertEqual('VAC',w.worksheets[0]['BC3'].value);w.close()
    def test_notification_one_per_property_with_all_legal_lots(self):
        path=self.save(self.split_parent());w=load_workbook(path,read_only=False,data_only=False)
        from wizards.mls.mailout_engine import read_owner_rows_from_step1
        owners=read_owner_rows_from_step1(w);w.close();self.assertEqual(1,len(owners));self.assertIn('LT 1',owners[0]['Legal_Dscr_1']);self.assertIn('LT 2',owners[0]['Legal_Dscr_1'])
    def test_existing_zero_area_excluded_row_is_preserved(self):
        source=self.save(parcel());rows,meta=workbook_rows(source)
        rows[0].update(land_area=0,land_use='EXC')
        out=self.folder/'excluded.xlsx';reviewed_workbook_copy(source,out,rows,meta,'mls')
        w=load_workbook(out,read_only=True,data_only=False)
        self.assertEqual(0,w.worksheets[0]['AO2'].value);self.assertEqual('EXC',w.worksheets[0]['BC2'].value);w.close()
    def record(self):return {'reviewed':True,'source':'Test plat','crs':'EPSG:2278','lots':[{'id':'a','geometry':mapping(box(0,0,50,60))},{'id':'b','geometry':mapping(box(50,0,100,60))}]}
    def test_geometry_identity_coverage_overlap_and_no_guess(self):
        r=self.record();parent=box(0,0,100,60);self.assertEqual(2,len(geometry_lots(r,parent,['a','b'])))
        with self.assertRaises(ValueError):geometry_lots(r,parent,['a','c'])
        r['lots'][1]['geometry']=mapping(box(40,0,100,60))
        with self.assertRaises(ValueError):geometry_lots(r,parent,['a','b'])
        r=self.record();r['lots'][1]['geometry']=mapping(box(55,0,100,60))
        with self.assertRaises(ValueError):geometry_lots(r,parent,['a','b'])
        self.assertEqual([],geometry_lots(None,parent))
    def test_draw_actual_split_preserves_union(self):
        parent=box(0,0,100,60);parts=split_geometry([parent],[(30,-1),(30,61)])
        self.assertEqual([1800,4200],sorted(p.area for p in parts));self.assertLess(parent.symmetric_difference(unary_union(parts)).area,.001)
    def test_tax_overlay_only_adds_interior(self):
        import fitz
        d={'case':{'lot_metadata':{'groups':{'123':self.record()}}},'matched':[{'property':{'account':'123','address':'Test','lot_rows':[{'id':'a'},{'id':'b'}]},'feature':{'geometry':{'rings':[[[0,0],[100,0],[100,60],[0,60],[0,0]]]}}}]}
        doc=fitz.open();doc.new_page(width=792,height=612);pdf=doc.tobytes();doc.close()
        result,notes=overlay_tax_lots(pdf,d,{},[-10,-30,110,90]);self.assertFalse(notes)
        with fitz.open(stream=result,filetype='pdf') as doc:self.assertGreater(len(doc[0].get_drawings()),0)
        unchanged,_=overlay_tax_lots(pdf,d,{'show_platted_lot_lines':False},[-10,-30,110,90]);self.assertEqual(pdf,unchanged)
    def test_roster_refresh_cache_warning_and_multiple_districts(self):
        html='<html>'+''.join(f'<h2>District {d}</h2><p>TEST MEMBER {d}</p>' for d in 'ABCDEFGHIJK')+'</html>'
        roster=parse_members(html);self.assertEqual('Test Member H',roster['H'])
        result=add_members({'districts':['C','H'],'warnings':[]},self.folder,lambda url:html)
        self.assertIn('District H: Test Member H',result['member_value'])
        def failure(url):raise OSError('offline')
        cached=add_members({'districts':['H'],'warnings':[]},self.folder,failure)
        self.assertEqual('Test Member H',cached['member_value']);self.assertIn('checked',cached['warnings'][0])
        with self.assertRaises(ValueError):parse_members('<h2>District C</h2><p>Only Member</p>')
    def test_council_copy_fills_both_fields(self):
        source=self.save(parcel());out=self.folder/'council.xlsx'
        write_case_district_copy(source,out,{'value':'H','member_value':'Test Member'})
        w=load_workbook(out,read_only=True,data_only=True);ci={r[0]:r[1] for r in w['Case Info'].iter_rows(max_col=2,values_only=True)};w.close()
        self.assertEqual(('H','Test Member'),(ci['Council_District'],ci['Council_Member']))
    def test_pc_divisions_preserve_other_reviewed_properties(self):
        from wizards.mls.pc_map import expand_pc_lots
        original={'matched':[{},{}],'pc_facts':{'rows':[]},'pc_protested_rows':[19]}
        expanded={'matched':[{'platted_lot':True},{'platted_lot':True},{'pc':{'land_use':'VAC','signed':True,'protested':True}}]}
        joined={'matched':[{'pc':{'row':17}},{'pc':{'row':18}},{'pc':{'land_use':'SFR','signed':False}}]}
        with patch('core.lot_divisions.expand_map_lots',return_value=expanded),patch('wizards.mls.pc_data.attach_pc_rows',return_value=joined) as attach:
            result=expand_pc_lots(original,{})
        self.assertEqual(17,result['matched'][0]['pc']['row'])
        self.assertEqual(expanded['matched'][2]['pc'],result['matched'][2]['pc'])
        attach.assert_called_once_with(expanded,original['pc_facts'],[19])

def main():return 0 if unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(LotsTests)).wasSuccessful() else 1
if __name__=='__main__':raise SystemExit(main())
