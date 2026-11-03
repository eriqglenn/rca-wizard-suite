"""V35 equal lot allocation and automatic map-to-Case-Info regressions."""
import copy,io,tempfile,unittest,zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from lxml import etree as ET
from openpyxl import load_workbook
from core.application_import import save_calculation,sheet_paths,set_cell
from core.automatic_case import ensure_lot_rows,save_map_council,load_map_data,_commit
from core.boundary_map import fingerprint
from core.hcad_identity import collect_hcad_records
from core.lot_divisions import proposed_lots,expand_parcels,workbook_rows,read_metadata
from core.council_district import ROSTER_URL
from tools.test_application_import import fields,parcel

KEYS=('mls','mbl','mls_renewal','mbl_renewal')
ROSTER='<html>'+''.join(f'<h2>District {d}</h2><p>TEST MEMBER {d}</p>' for d in 'ABCDEFGHIJK')+'</html>'


class SimpleFlowTests(unittest.TestCase):
    def setUp(self):
        from tools.test_designation_lookup import no_lookup
        external=patch('core.designation_lookup.lookup',side_effect=no_lookup);external.start();self.addCleanup(external.stop)
        self.temp=tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup);self.folder=Path(self.temp.name)

    def parent(self,area=15000):
        row=parcel(area=area);row.update(land_area=area,legal_description='LTS 1, 2, 3 BLK 69',
            original_lot_area=12000,building_line=25,original_building_line=30)
        return row

    def save(self,key='mls',row=None):
        f={**fields(),'renewal_minimum':'3000' if key.startswith('mls') else '25'}
        f['application_number']='910REN' if key.endswith('renewal') else '910'
        return save_calculation(self.folder/(key+'.xlsx'),f,{'parcels':[row or self.parent()],
            'warnings':[],'sources':[]},{},key,reviewed=True)

    def case_info(self,path):
        book=load_workbook(path,read_only=True,data_only=True)
        try:return dict(book['Case Info'].iter_rows(max_col=2,values_only=True))
        finally:book.close()

    def make_unsplit(self,key):
        row=self.parent();row['legal_description']='LT 1 BLK 69'
        source=self.save(key,row)
        with zipfile.ZipFile(source) as z:
            parts={n:z.read(n) for n in z.namelist()};sheet=next(p for n,p in sheet_paths(z).items() if n.startswith('Step 1'))
        xml=ET.fromstring(parts[sheet]);set_cell(xml,'AD2','LTS 1, 2, 3 BLK 69');parts[sheet]=ET.tostring(xml)
        with zipfile.ZipFile(source,'w',zipfile.ZIP_DEFLATED) as z:
            for n,data in parts.items():z.writestr(n,data)
        return source

    def test_equal_allocation_preserves_area_and_measurements(self):
        row=self.parent();lots=proposed_lots(row)
        self.assertEqual([5000]*3,[l['area'] for l in lots])
        self.assertEqual([4000]*3,[l['original_lot_area'] for l in lots])
        self.assertEqual([25]*3,[l['building_line'] for l in lots])
        self.assertEqual([30]*3,[l['original_building_line'] for l in lots])
        self.assertTrue(all(l['allocation']=='equal' for l in lots))
        for area in (10000,0,None):
            with self.subTest(area=area):
                expanded=expand_parcels([self.parent(area)])
                self.assertEqual(3,len(expanded))
                if area is None:self.assertTrue(all(p['land_area'] is None for p in expanded))
                else:self.assertAlmostEqual(area,sum(p['land_area'] for p in expanded),places=7)
        row['legal_description']='W 25 FT OF LTS 1 & 2 BLK 69'
        self.assertEqual([],proposed_lots(row))

    def test_new_calculation_all_four_types_splits_without_review(self):
        for key in KEYS:
            with self.subTest(wizard=key):
                path=self.save(key);rows,_=workbook_rows(path)
                self.assertEqual([5000]*3,[r['land_area'] for r in rows])
                self.assertEqual(1,len(collect_hcad_records(path)))
                book=load_workbook(path,read_only=True,data_only=False)
                try:
                    self.assertEqual(56,book.worksheets[0].max_column)
                    self.assertEqual(['LT 1 BLK 69','LT 2 BLK 69','LT 3 BLK 69'],[book.worksheets[0][f'AD{n}'].value for n in (2,3,4)])
                    if key=='mls_renewal':self.assertEqual([4000]*3,[book.worksheets[1][f'D{n}'].value for n in (15,16,17)])
                    if key=='mbl_renewal':
                        self.assertEqual([30]*3,[book.worksheets[1][f'C{n}'].value for n in (15,16,17)])
                        self.assertEqual([25]*3,[book.worksheets[1][f'D{n}'].value for n in (15,16,17)])
                    self.assertEqual('Equal division',read_metadata(book)['rows'][0]['area_source']['method'])
                finally:book.close()

    def test_existing_calculations_split_once_and_preserve_step2(self):
        for key in KEYS:
            with self.subTest(wizard=key):
                source=self.make_unsplit(key);before=source.read_bytes()
                result=ensure_lot_rows(source,key,self.folder)
                self.assertEqual(str(source),result['path']);self.assertEqual(1,result['split_properties'])
                rows,_=workbook_rows(source);self.assertEqual([5000]*3,[r['land_area'] for r in rows])
                if key=='mls_renewal':self.assertEqual([4000]*3,[r['original_lot_area'] for r in rows])
                if key=='mbl_renewal':self.assertEqual([25]*3,[r['building_line'] for r in rows])
                with zipfile.ZipFile(io.BytesIO(before)) as old,zipfile.ZipFile(source) as new:
                    for n in old.namelist():
                        if n.startswith(('xl/media/','xl/drawings/','xl/styles')):self.assertEqual(old.read(n),new.read(n),n)
                after=source.read_bytes();again=ensure_lot_rows(source,key,self.folder)
                self.assertEqual(0,again['split_properties']);self.assertEqual(after,source.read_bytes())
                backups=list(source.with_name(source.stem+'_sources').glob('Workbook backups/*.xlsx'))
                self.assertEqual(1,len(backups));self.assertEqual(before,backups[0].read_bytes())

    def test_council_update_only_changes_case_info_and_is_idempotent(self):
        source=self.save();before=source.read_bytes();result={'value':'H','member_value':'Test Member H'}
        path,notes=save_map_council(source,result,self.folder,fingerprint(source))
        self.assertEqual(source,path);self.assertFalse(notes)
        ci=self.case_info(source);self.assertEqual(('H','Test Member H'),(ci['Council_District'],ci['Council_Member']))
        with zipfile.ZipFile(io.BytesIO(before)) as old,zipfile.ZipFile(source) as new:
            case=sheet_paths(old)['Case Info']
            self.assertEqual(old.namelist(),new.namelist())
            for name in old.namelist():
                if name!=case:self.assertEqual(old.read(name),new.read(name),name)
        after=source.read_bytes();save_map_council(source,result,self.folder)
        self.assertEqual(after,source.read_bytes())

    def test_locked_workbook_uses_copy_and_stale_source_is_not_overwritten(self):
        source=self.save();before=source.read_bytes()
        with patch('core.automatic_case.os.replace',side_effect=PermissionError('Excel is open')):
            result,notes=save_map_council(source,{'value':'H','member_value':'Test Member H'},self.folder)
        self.assertNotEqual(source,result);self.assertEqual(before,source.read_bytes());self.assertTrue(notes)
        self.assertEqual('H',self.case_info(result)['Council_District'])
        with self.assertRaisesRegex(ValueError,'changed'):_commit(source,b'before someone saved',before,self.folder)
        self.assertEqual(before,source.read_bytes())

    def test_failed_member_lookup_preserves_only_same_district_name(self):
        source=self.save();save_map_council(source,{'value':'H','member_value':'Test Member H'})
        save_map_council(source,{'value':'H','member_value':''})
        self.assertEqual('Test Member H',self.case_info(source)['Council_Member'])
        save_map_council(source,{'value':'C','member_value':''})
        self.assertFalse(self.case_info(source)['Council_Member'])

    def test_map_creation_updates_all_four_case_sheets_and_hash(self):
        from core.council_district import add_members
        for key in KEYS:
            with self.subTest(wizard=key):
                source=self.save(key)
                data={'case':{'workbook_hash':fingerprint(source)},'matched':[{'feature':{'geometry':{'rings':[[[0,0],[100,0],[100,100],[0,100],[0,0]]]},'attributes':{'HCAD_NUM':'0350820240027'}}}],
                      'warnings':[],'sources':[]}
                def query(client,*args,**kwargs):
                    return [{'attributes':{'DISTRICT':'H'},'geometry':{'rings':[[[-10,-10],[110,-10],[110,110],[-10,110],[-10,-10]]]}}]
                with patch('core.automatic_case.WorkbookReadCache.resolve',return_value=str(source)),patch('core.boundary_map.read_map_case',return_value=SimpleNamespace()),patch('core.boundary_map.fetch_map_data',return_value=copy.deepcopy(data)),patch('core.boundary_map.PublicMapClient.query',query),patch('core.council_district.add_members',side_effect=lambda result,cache:add_members(result,cache,lambda url:ROSTER)):
                    result=load_map_data(source,key,self.folder)
                self.assertTrue(result['council_saved']);self.assertEqual(1,len(result['matched']))
                self.assertEqual(fingerprint(source),result['case']['workbook_hash'])
                self.assertEqual('Test Member H',self.case_info(source)['Council_Member'])
                self.assertEqual(ROSTER_URL,result['council']['member_source'])

    def test_council_service_failure_does_not_block_the_map(self):
        source=self.save();before=source.read_bytes()
        data={'case':{'workbook_hash':fingerprint(source)},'matched':[{'feature':{}}],'warnings':[]}
        with patch('core.automatic_case.WorkbookReadCache.resolve',return_value=str(source)),patch('core.boundary_map.read_map_case',return_value=SimpleNamespace()),patch('core.boundary_map.fetch_map_data',return_value=data),patch('core.council_district.resolve_districts',side_effect=OSError('City service offline')):
            result=load_map_data(source,'mls',self.folder)
        self.assertEqual(1,len(result['matched']));self.assertEqual(before,source.read_bytes())
        self.assertIn('City service offline',result['warnings'][-1]);self.assertNotIn('council_saved',result)

    def test_city_district_touching_rings_preserve_both_parts(self):
        from core.council_district import district_polygon
        from shapely.geometry import Point
        feature={'geometry':{'rings':[[[0,0],[10,0],[10,10],[0,10],[0,0],[-10,0],[-10,-10],[0,-10],[0,0]]]}}
        shape=district_polygon(feature)
        self.assertTrue(shape.is_valid);self.assertEqual(200,shape.area)
        self.assertTrue(shape.covers(Point(5,5)));self.assertTrue(shape.covers(Point(-5,-5)))


def main():
    return 0 if unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(SimpleFlowTests)).wasSuccessful() else 1


if __name__=='__main__':raise SystemExit(main())
