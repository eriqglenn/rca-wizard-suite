"""Cross-wizard account identity and PDF reuse checks; no browser/network calls."""
import importlib,io,os,tempfile,unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch,MagicMock
from openpyxl import Workbook
from core import hcad_downloader as hcad
from core.hcad_identity import clean_account

MODULES={'MLS':'wizards.mls.mailout_engine','MBL':'wizards.mbl.mbl_mailout_engine',
         'MLS Renewal':'wizards.mls_renewal.wizard','MBL Renewal':'wizards.mbl_renewal.wizard'}
A='0350820240027';B='0350820240028';C='0350820240029'

class HCADIdentityTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup);self.root=Path(self.temp.name)
    def workbook(self,rows):
        book=Workbook();sheet=book.active;sheet.title='Step 1 - Input Properties Info'
        sheet.append(['Site_addr_1','TAX_ID','ALT_TAX_ID','HCAD_NUM','HCAD_NUM_1','str_unit','Legal_Dscr_1','OWNER_MAILTO',
                      'Mail_Addr_1','Mail_Addr_2','Mail_City','Mail_State','Mail_Zip','Total_Land_Area','group_dscr'])
        for row in rows:sheet.append(row)
        info=book.create_sheet('Case Info');info.append(['Application_Number','123'])
        path=self.root/'fixture.xlsx';book.save(path);book.close();return path
    def exports(self,path,count,accounts):
        for key,module in MODULES.items():
            with self.subTest(wizard=key):
                filename,display_count=importlib.import_module(module).export_hcad_address_list(path,self.root/key)
                items=hcad.load_search_items(filename)
                self.assertEqual(count,display_count);self.assertEqual(count,len(items))
                self.assertEqual(set(accounts),{i.account for i in items})
    def test_accounts_normalize_without_inventing_invalid_numbers(self):
        for value in (A,A[1:],int(A),float(int(A)),str(int(A))+'.0','3.50820240027E+11'):
            self.assertEqual(A,clean_account(value))
        for value in ('123','12345678901234',float('nan'),12.5,'1e1000000',-int(A)):self.assertEqual('',clean_account(value))
    def test_all_four_exports_count_accounts_not_lots_or_owners(self):
        rows=[['100 TEST ST',A],['100 TEST ST LOT 2',A[1:]],['100 TEST ST LOT 3',str(int(A))+'.0'],['100 TEST ST',B],['',C]]
        rows=[row+[None]*(7-len(row))+['SAME OWNER'] for row in rows]
        self.exports(self.workbook(rows),3,[A,B,C])
    def test_all_four_exports_support_alternate_account_columns(self):
        self.exports(self.workbook([['100 TEST ST','invalid',A],['100 TEST ST',None,None,B],['100 TEST ST',None,None,None,C]]),3,[A,B,C])
    def test_address_fallback_is_used_only_without_a_valid_account(self):
        self.exports(self.workbook([['100 TEST ST',None],['  100   test st ',None],['100 TEST ST',A]]),2,['',A])
    def test_file_and_manual_batches_share_duplicate_protection(self):
        path=self.root/'items.txt';path.write_text(f'100 TEST ST\t{A}\n100 TEST ST LOT 2\t{A[1:]}\n100 TEST ST\t{B}\n')
        self.assertEqual([A,B],[i.account for i in hcad.load_search_items(path)])
        raw=[hcad.SearchItem('100 TEST ST',A),hcad.SearchItem('100 TEST ST LOT 2',A[1:]),hcad.SearchItem('100 TEST ST',B)]
        for resume in (True,False):
            with patch.object(hcad,'download_hcad_record',side_effect=lambda page,item,resume:hcad.DownloadResult(item,'Success',str(self.root/(item.account+'.pdf')))) as download,patch.object(hcad,'wait_for_hcad_ready'),patch.object(hcad.time,'sleep'),redirect_stdout(io.StringIO()):
                results=hcad.run_batch(MagicMock(),raw,0,0,resume)
            self.assertEqual(2,download.call_count);self.assertEqual([A,B],[r.item.account for r in results])
    def pdf(self,name,account):
        import fitz
        path=self.root/name
        with fitz.open() as doc:
            page=doc.new_page();page.insert_text((72,72),f'Account Number: {account}\nLand Area: 6,000 SF');doc.save(path)
        return path
    def test_existing_account_record_wins_over_shared_address(self):
        self.pdf('HCAD_100_TEST_ST_'+B+'.pdf',B)
        expected=self.pdf('HCAD_OLD_ADDRESS_'+A+'.pdf',A)
        with patch.object(hcad,'SAVE_FOLDER',self.root),patch.object(hcad,'perform_search') as search,redirect_stdout(io.StringIO()):
            result=hcad.download_hcad_record(None,hcad.SearchItem('100 TEST ST',A))
        self.assertEqual('Skipped',result.status);self.assertEqual(expected,Path(result.pdf_path));search.assert_not_called()
    def test_misnamed_or_unreadable_pdf_does_not_hide_valid_pdf(self):
        self.pdf('HCAD_A_WRONG_'+A+'.pdf',B)
        (self.root/('HCAD_B_CORRUPT_'+A+'.pdf')).write_bytes(b'invalid test PDF')
        (self.root/('HCAD_C_EMPTY_'+A+'.pdf')).touch()
        expected=self.pdf('HCAD_D_GOOD_'+A+'.pdf',A)
        with patch.object(hcad,'SAVE_FOLDER',self.root):
            self.assertEqual(expected,hcad.find_existing_pdf(hcad.SearchItem('100 TEST ST',A)))
    def test_wrong_account_is_not_reused(self):
        self.pdf('HCAD_100_TEST_ST_'+A+'.pdf',B)
        with patch.object(hcad,'SAVE_FOLDER',self.root):self.assertIsNone(hcad.find_existing_pdf(hcad.SearchItem('100 TEST ST',A)))
    def test_all_four_ui_export_messages_identify_unique_properties(self):
        for key,cls in [('mls','MailoutWizard'),('mbl','MBLWizard'),('mls_renewal','MLSRenewalWizard'),('mbl_renewal','MBLRenewalWizard')]:
            module=importlib.import_module('wizards.'+key+'.wizard')
            constants=getattr(module,cls)._export_hcad_list.__code__.co_consts
            self.assertTrue(any('unique HCAD property record(s)' in value for value in constants if isinstance(value,str)),key)
    def test_actual_187_case_across_all_four_exporters(self):
        path=os.environ.get('RCA_HCAD_CASE')
        if not path:self.skipTest('Set RCA_HCAD_CASE for the actual 187REN case.')
        from core.lot_divisions import workbook_rows
        rows,_=workbook_rows(path);self.assertEqual(45,len(rows));accounts={r['account'] for r in rows};self.assertEqual(30,len(accounts))
        self.exports(path,30,accounts)

def main():
    return 0 if unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(HCADIdentityTests)).wasSuccessful() else 1
if __name__=='__main__':raise SystemExit(main())
