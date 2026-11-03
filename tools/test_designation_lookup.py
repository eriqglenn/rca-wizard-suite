"""Renewal lookup: boundary identity, provenance, dates, safe saves and actual UI flow."""
import copy
import io
import json
import tempfile
import unittest
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import load_workbook
from core import designation_lookup as dl
from core.application_import import save_calculation, sheet_paths
from core.automatic_case import load_map_data
from core.boundary_map import fingerprint
from tools.test_application_import import fields, parcel


def no_lookup(*args, **kwargs):
    return {'values':{}, 'warnings':[], 'sources':[], 'candidates':[], 'checked_utc':'2026-09-18T12:00:00+00:00'}


def feature(x0=0,x1=100,number='2006-1019',app='187',status='Approved',family='mls'):
    attrs={'OBJECTID':1,'App_Num':app,'Status':status,'Type':'NEW','Expiration_Date':1791590400000,
           'Effective_Date':1160438400000,'STARTDATE':1159920000000,'START_DATE':1159920000000,
           'ORDINANCE':number,'ORDINANCE_':number}
    return {'attributes':attrs,'geometry':{'rings':[[[x0,0],[x1,0],[x1,60],[x0,60],[x0,0]]]}}


class Repository:
    def __init__(self, records=None, passage='10/04/2006'):
        self.records=records if records is not None else [feature()]
        self.date=passage;self.warnings=[];self.sources=[]
        self.client=SimpleNamespace(check_cancelled=lambda:None)
    def designations(self,family):return self.records
    def status_records(self,family):return [{'ordinance':'2006-1019','passage':self.date,'source':'https://www.houstontx.gov/test.xlsx | row 2'}]
    def passage(self,number,hint):return None


class DesignationTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup);self.folder=Path(self.tmp.name)
    def run_lookup(self,repo=None,fields_=None,wizard='mls_renewal'):
        return dl.lookup([parcel()['feature']],wizard,fields_ or {'application_number':'187REN'},self.folder,repository=repo or Repository())
    def save(self,wizard='mls_renewal',result=None,extra=None):
        f={**fields(),'application_number':'187REN','renewal_minimum':'3000' if wizard.startswith('mls') else '19',**(extra or {})}
        path=self.folder/(wizard+'.xlsx')
        return save_calculation(path,f,{'parcels':[parcel()],'warnings':[],'sources':[]},{'renewal_lookup':result} if result else {},wizard,reviewed=True)
    def test_dates_preserve_utc_calendar_and_distinct_meanings(self):
        result=self.run_lookup()
        self.assertEqual({'original_ordinance_no':'2006-1019','original_passage_date':'10/04/2006','original_effective_date':'10/10/2006','original_expiration_date':'10/10/2026'},result['values'])
        self.assertEqual('2006-0007',dl.ordinance('Ordinance No. 2006-7'))
        self.assertEqual('09/19/2006',dl.date_text(1158624000000))
    def test_mbl_start_is_not_assumed_to_be_effective_or_passage(self):
        result=self.run_lookup(Repository(passage=''),wizard='mbl_renewal')
        self.assertNotIn('original_effective_date',result['values']);self.assertNotIn('original_passage_date',result['values'])
        self.assertIn('original_expiration_date',result['values'])
    def test_shared_edge_outlier_and_partial_boundary_do_not_fill(self):
        for records in ([feature(100,200)],[feature(0,50)]):
            self.assertFalse(self.run_lookup(Repository(records))['values'])
        features=[parcel()['feature'],feature(1000,1100)]
        result=dl.lookup(features,'mls_renewal',{},self.folder,repository=Repository())
        self.assertFalse(result['values'])
    def test_split_geometry_for_same_ordinance_is_combined(self):
        result=self.run_lookup(Repository([feature(0,50),feature(50,100)]))
        self.assertEqual('2006-1019',result['values']['original_ordinance_no'])
    def test_ambiguous_boundaries_and_explicit_wrong_ordinance_keep_blanks(self):
        a=feature();b=feature(number='2007-0010',app='900')
        result=self.run_lookup(Repository([a,b]));self.assertFalse(result['values']);self.assertTrue(result['candidates'])
        result=self.run_lookup(Repository(),{'original_ordinance_no':'2006-0001'})
        self.assertFalse(result['values'])
    def test_latest_prior_renewal_same_lineage_and_footprint(self):
        a=feature();b=feature(number='2020-0010',app='187REN');b['attributes'].update(Type='REN',Effective_Date=1578441600000)
        result=self.run_lookup(Repository([a,b]),{'application_number':'187REN','date_received':'01/01/2026'})
        self.assertEqual('2020-0010',result['values']['original_ordinance_no'])
        b['attributes']['Effective_Date']=1791590400000
        result=self.run_lookup(Repository([a,b]),{'application_number':'187REN','date_received':'01/01/2026'})
        self.assertEqual('2006-1019',result['values']['original_ordinance_no'])
    def test_unapproved_and_rescinded_records_do_not_supply_values(self):
        for state in ('Pending','Denied','Rescinded'):
            self.assertFalse(self.run_lookup(Repository([feature(status=state)]))['values'])
    def test_conflicting_published_dates_are_not_filled(self):
        a=feature();b=feature();b['attributes']['Effective_Date']=1160524800000
        result=self.run_lookup(Repository([a,b]))
        self.assertNotIn('original_effective_date',result['values']);self.assertTrue(any('disagree' in s for s in result['warnings']))
        a['attributes']['Expiration_Date']=1_000_000_000_000
        self.assertNotIn('original_expiration_date',self.run_lookup(Repository([a]))['values'])
    def test_existing_fields_and_wrong_identity_are_never_overwritten(self):
        result=self.run_lookup();values,filled,notes=dl.merge_fields({'original_effective_date':'10/11/2006'},result)
        self.assertEqual('10/11/2006',values['original_effective_date']);self.assertTrue(notes)
        values,filled,notes=dl.merge_fields({'original_ordinance_no':'2001-123'},result)
        self.assertEqual([],filled);self.assertNotIn('original_expiration_date',values)
    def test_adoption_parser_rejects_historical_reference_and_ambiguous_date(self):
        self.assertEqual('10/04/2006',dl.adopted_date('WEDNESDAY - OCTOBER 4, 2006\nORDINANCE 2006-1019 ADOPTED','2006-1019'))
        self.assertEqual('',dl.adopted_date('WEDNESDAY - OCTOBER 4, 2026\nrenewing Ordinance No. 2006-1019','2006-1019'))
        self.assertEqual('',dl.adopted_date('WEDNESDAY - OCTOBER 4, 2006\nORDINANCE 2006-1019 ADOPTED\nTHURSDAY - OCTOBER 5, 2006\nORDINANCE 2006-1019 ADOPTED','2006-1019'))
        minutes='TUESDAY - OCTOBER 29, 2024\nPublic speakers. The Houston City Council reconvened at 9:00 a.m., Wednesday, October 30, 2024.\nHEARING DATE - WEDNESDAY - NOVEMBER 13, 2024\nORDINANCE 2024-0819 ADOPTED'
        self.assertEqual('10/30/2024',dl.adopted_date(minutes,'2024-0819'))
    def test_persisted_inputs_dates_provenance_and_zip_parts(self):
        result=self.run_lookup();path=self.save();before=path.read_bytes()
        updated,notes=dl.save_workbook(path,result,self.folder,fingerprint(path))
        self.assertEqual(path,updated);self.assertIn('Filled blank',notes[0])
        book=load_workbook(path,read_only=True,data_only=True)
        try:
            values=dict(book['Case Info'].iter_rows(max_col=2,values_only=True))
            self.assertEqual('2006-1019',values['Original_Ordinance_No'])
            self.assertEqual(datetime(2006,10,4),values['Original_Passage_Date'])
            self.assertIn('houstontx.gov',values['Renewal_Lookup_Source'])
        finally:book.close()
        with zipfile.ZipFile(io.BytesIO(before)) as old,zipfile.ZipFile(path) as new:
            case=sheet_paths(old)['Case Info']
            self.assertEqual(old.namelist(),new.namelist())
            for name in old.namelist():
                if name!=case:self.assertEqual(old.read(name),new.read(name),name)
        after=path.read_bytes();again=self.run_lookup();dl.save_workbook(path,again,self.folder)
        self.assertEqual(after,path.read_bytes())
    def test_import_saves_lookup_for_both_renewals_and_manual_passage(self):
        for wizard in ('mls_renewal','mbl_renewal'):
            path=self.save(wizard,self.run_lookup(wizard=wizard))
            info=dl.workbook_fields(path)
            self.assertEqual('2006-1019',info['original_ordinance_no'])
            self.assertEqual(datetime(2006,10,4),info['original_passage_date'])
        path=self.folder/'mls_renewal.xlsx';path.unlink()
        # Reuse a different case destination because the original sources are retained.
        self.folder=self.folder/'manual';self.folder.mkdir()
        path=self.save(extra={'original_passage_date':'09/01/2006'})
        self.assertEqual(datetime(2006,9,1),dl.workbook_fields(path)['original_passage_date'])
    def test_network_failure_is_a_warning_and_does_not_block_workbook(self):
        repo=Repository();repo.designations=lambda family:(_ for _ in ()).throw(OSError('offline'))
        result=self.run_lookup(repo);self.assertFalse(result['values']);self.assertIn('offline',result['warnings'][0])
        path=self.save(result=result);self.assertTrue(path.is_file())
    def test_locked_workbook_uses_copy_and_changed_file_is_protected(self):
        path=self.save();before=path.read_bytes()
        with patch('core.automatic_case.os.replace',side_effect=PermissionError('Excel is open')):
            updated,notes=dl.save_workbook(path,self.run_lookup(),self.folder)
        self.assertNotEqual(path,updated);self.assertEqual(before,path.read_bytes())
        with self.assertRaisesRegex(ValueError,'changed'):dl.save_workbook(path,self.run_lookup(),self.folder,'wrong')
    def test_map_lookup_runs_after_council_save_and_hash_stays_current(self):
        for wizard in ('mls_renewal','mbl_renewal'):
            path=self.save(wizard);result=self.run_lookup(wizard=wizard)
            data={'case':{'workbook_hash':fingerprint(path)},'matched':[{'feature':parcel()['feature']}],'warnings':[],'sources':[]}
            council={'value':'H','districts':['H'],'member_value':'Test Member','warnings':[]}
            with patch('core.automatic_case.WorkbookReadCache.resolve',return_value=str(path)),patch('core.boundary_map.read_map_case',return_value=SimpleNamespace()),patch('core.boundary_map.fetch_map_data',return_value=data),patch('core.council_district.resolve_districts',return_value=council),patch('core.council_district.add_members',side_effect=lambda r,c:r),patch.object(dl,'lookup',return_value=result):
                mapped=load_map_data(path,wizard,self.folder)
            self.assertEqual(fingerprint(path),mapped['case']['workbook_hash'])
            info=dl.workbook_fields(path);self.assertEqual('H',info['council_district']);self.assertEqual('2006-1019',info['original_ordinance_no'])
    def test_new_applications_do_not_request_previous_ordinance(self):
        with patch.object(dl,'CityRecords',side_effect=AssertionError('no lookup')):
            for wizard in ('mls','mbl'):self.assertEqual({},dl.lookup([],wizard,{},self.folder)['values'])

    def test_import_save_button_fills_renewal_fields_and_retains_manual_date(self):
        import importlib,time,tkinter as tk
        from core import session
        from core.application_import_ui import ApplicationImport
        root=tk.Tk();root.withdraw();self.addCleanup(root.destroy)
        with patch.object(session,'SETTINGS_PATH',self.folder/'session.json'),patch.object(session,'load',return_value={}):
            for wizard,cls in [('mls_renewal','MLSRenewalWizard'),('mbl_renewal','MBLRenewalWizard')]:
                owner=getattr(importlib.import_module('wizards.'+wizard+'.wizard'),cls)(root)
                owner.files['out']=str(self.folder)
                dialog=ApplicationImport(owner);dialog.withdraw()
                included=parcel();outside=parcel(number=2016,account='0200650000016');outside['include']=False
                dialog.data={'parcels':[included,outside],'fields':fields(),'roads':[],'warnings':[],'sources':[]};dialog.application={}
                for k,v in fields().items():
                    if k in dialog.fields:dialog.fields[k].set(v)
                dialog.fields['renewal_minimum'].set('3000' if wizard.startswith('mls') else '19')
                dialog.fields['original_passage_date'].set('09/01/2006')
                dialog.case_reviewed.set(True);dialog.parcels_reviewed.set(True)
                destination=self.folder/(wizard+'_ui.xlsx');result=self.run_lookup(wizard=wizard)
                def fail(*args,**kwargs):raise AssertionError(str(args))
                with patch('tkinter.filedialog.asksaveasfilename',return_value=str(destination)),patch('tkinter.messagebox.showinfo'),patch('tkinter.messagebox.showerror',side_effect=fail),patch.object(dl,'lookup',return_value=result) as lookup:
                    dialog.save_button.invoke();deadline=time.monotonic()+30
                    while not dialog.closed and time.monotonic()<deadline:root.update();time.sleep(.02)
                    self.assertTrue(dialog.closed,dialog.status.get())
                    self.assertEqual([included['feature']],lookup.call_args.args[0])
                info=dl.workbook_fields(destination)
                self.assertEqual('2006-1019',info['original_ordinance_no'])
                self.assertEqual(datetime(2006,9,1),info['original_passage_date'])
                self.assertIn('kept',info['renewal_lookup_notes'])
                self.assertEqual(str(destination),owner.files['workbook']);owner.destroy()

    def test_public_status_parser_and_cache_fallback(self):
        import os
        # An existing City-format fixture has no approval-date heading; no date is invented.
        path=self.save();self.assertEqual([],dl.status_rows(path.read_bytes(),'mls','https://www.houstontx.gov/status.xlsx'))
        repo=dl.CityRecords(self.folder/'cache')
        with patch.object(repo.client,'query',return_value=[feature()]) as query:
            first=repo.designations('mls');second=repo.designations('mls');self.assertEqual(first,second)
            query.assert_called_once_with(dl.SERVICES['mls'],fields='*')
        repo.refresh=True
        with patch.object(repo.client,'query',side_effect=OSError('offline')):
            self.assertEqual(first,repo.designations('mls'));self.assertTrue(repo.warnings)
        with self.assertRaisesRegex(ValueError,'official City'):repo.blob('https://other.example/file.xlsx')


def main():
    return 0 if unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(DesignationTests)).wasSuccessful() else 1


if __name__=='__main__':raise SystemExit(main())
