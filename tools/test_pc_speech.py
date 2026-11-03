"""PC speech data, missing-fact, Word and real Tk integration regression tests."""
import copy
from pathlib import Path
import tempfile
import time
import tkinter as tk
from tkinter import ttk
import unittest
from unittest.mock import patch
from openpyxl import load_workbook
from docx import Document
from core import session
from core.boundary_map import fingerprint
from tools.test_pc_report import case_fixture
from wizards.mls.pc_data import read_pc_facts, load_preferences, save_preferences
from wizards.mls.pc_speech import (defaults, read_speech_facts, speech_context, speech_text,
                                   generate_speech, fill_speech)

def options():
    return dict(agenda='12',presenter='Taylor Planner',neighborhood='Example',district='C',
                super_neighborhood='Example Central',subdivision='Example Addition',plat_year='1926',
                earliest_year='1928',protest_count='0',notice_mailed=True,recommendation='Approval')

class PCSpeechTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup)
        self.folder=Path(self.temp.name);self.path,self.facts,_=case_fixture(self.folder)

    def test_denominators_ties_excluded_and_split_rows(self):
        # Same owner/account may represent several individual legal lots; every row counts.
        ctx,notes=speech_context(self.facts,options());text=speech_text(ctx)
        self.assertFalse(notes)
        self.assertIn('3 single-family residential lots, representing 75%',text)
        self.assertIn('1 commercial lot, representing 25% of the eligible application land area',text)
        self.assertIn('owners of 2 of the 4 eligible lots, representing 35%',text)
        self.assertIn('at least 5,000 square feet exists on 3 lots',text)
        self.assertIn('1 excluded lot is outside',text)
        self.assertIn(self.facts['boundary'],text)
        self.assertNotIn('X,XXX',text);self.assertNotIn('19XX',text);self.assertNotIn('{{',text)
        # Three 5,000-sq-ft lots on a 15,000-sq-ft property all count and carry support.
        f=copy.deepcopy(self.facts)
        f['rows']=[dict(row=i,description='LOT '+str(i),land_use='SFR',area=5000,signed=True,support_known=True,account='same') for i in (1,2,3)]
        f.update(excluded_count=0,eligible_count=3)
        ctx,_=speech_context(f,options())
        self.assertIn('3 single-family residential lots, representing 100%',ctx['land_use_sentence'])
        self.assertIn('owners of 3 of the 3 lots, representing 100%',ctx['support_sentence'])
        self.assertEqual('3',ctx['qualifying_count'])

    def test_missing_facts_are_visible_not_invented(self):
        opts=options();opts.update(agenda='',earliest_year='',protest_count='',notice_mailed=False)
        ctx,notes=speech_context(self.facts,opts)
        self.assertGreaterEqual(len(notes),4)
        self.assertIn('[Enter agenda item]',speech_text(ctx))
        self.assertIn('[Enter earliest house year]',speech_text(ctx))
        self.assertIn('[Confirm',ctx['notice_sentence']);self.assertIn('[Confirm',ctx['protest_sentence'])
        self.assertNotIn('No timely written protests',ctx['protest_sentence'])
        f=copy.deepcopy(self.facts);f['rows'][0]['area']=None
        ctx,notes=speech_context(f,options())
        self.assertIn('[Review',ctx['land_use_sentence']);self.assertIn('[Review',ctx['support_sentence'])
        f=copy.deepcopy(self.facts);f['rows'][0]['support_known']=False
        ctx,notes=speech_context(f,options());self.assertIn('[Review',ctx['support_sentence'])

    def test_other_uses_and_protest_grammar(self):
        f=copy.deepcopy(self.facts);f['rows'][0]['land_use']='MF';f['rows'][1]['land_use']='VAC'
        ctx,_=speech_context(f,options());self.assertIn('1 multifamily residential lot and 1 vacant lot',ctx['land_use_sentence'])
        self.assertEqual('No timely written protests were received by staff.',ctx['protest_sentence'])
        for n,phrase in [('1','One timely written protest from a property owner was'),('2','2 timely written protests from property owners were')]:
            ctx,_=speech_context(f,{**options(),'protest_count':n});self.assertIn(phrase,ctx['protest_sentence'])
        for n in ('-1','1.5','unknown'):
            ctx,notes=speech_context(f,{**options(),'protest_count':n});self.assertIn('[Confirm',ctx['protest_sentence']);self.assertTrue(notes)

    def test_defaults_reuse_only_same_meaning_and_current_workbook(self):
        f=copy.deepcopy(self.facts);f['case_info'].update(Planner_Name='Planner A',Council_District='H',Neighborhood_Name='Area A',Super_Neighborhood='Area 23')
        prefs=dict(agenda='IX',plat_year='1931',protest_count=2,neighborhood='The neighborhood was platted in 1931.',notice_sent=True)
        opts=defaults(f,prefs)
        self.assertEqual(('Planner A','H','Area A','Area 23'),tuple(opts[k] for k in ('presenter','district','neighborhood','super_neighborhood')))
        self.assertEqual('IX',opts['agenda']);self.assertFalse(opts['notice_mailed'])
        save_preferences(self.path,self.folder,f['workbook_hash'],{'speech':options()})
        self.assertEqual({},load_preferences(self.path,self.folder,'different hash'))

    def test_full_year_and_subdivision_suffix(self):
        ctx,notes=speech_context(self.facts,{**options(),'agenda':'XX','plat_year':'2001','earliest_year':'2003','subdivision':'Example Subdivision'})
        text=speech_text(ctx);self.assertFalse(notes)
        self.assertIn('Item XX',text)  # Roman agenda numbers are valid entered data.
        self.assertIn('platted in 2001',text);self.assertIn('constructed in 2003',text)
        self.assertNotIn('Subdivision Subdivision',text)
        ctx,notes=speech_context(self.facts,{**options(),'earliest_year':'19XX'})
        self.assertIn('[Enter',ctx['earliest_year']);self.assertTrue(notes)

    def test_word_export_preserves_original_reading_format_and_workbook(self):
        before=self.path.read_bytes()
        result=generate_speech(self.path,self.facts,options(),self.folder,export_pdf=False)
        doc=Document(result['docx']);self.assertFalse(result['notes']);self.assertTrue(Path(result['folder'],'PC_Speech_Review.json').is_file())
        self.assertAlmostEqual(doc.sections[0].page_width.inches,8.5)
        self.assertAlmostEqual(doc.sections[0].left_margin.inches,1)
        self.assertEqual(14,doc.paragraphs[1].runs[0].font.size.pt)
        self.assertIn('STOP HERE',doc.paragraphs[5].text);self.assertEqual('Staff report: ',doc.paragraphs[6].text)
        self.assertFalse(any(r.font.color.rgb and str(r.font.color.rgb)=='EE0000' for p in doc.paragraphs for r in p.runs))
        self.assertEqual(before,self.path.read_bytes())
        draft=generate_speech(self.path,self.facts,{},self.folder,export_pdf=False)
        self.assertTrue(draft['docx'].endswith('_DRAFT.docx'));self.assertTrue(draft['notes'])

    def test_changed_workbook_and_non_mls_are_rejected(self):
        wb=load_workbook(self.path);wb.worksheets[1]['A1']='MINIMUM BUILDING LINE';wb.save(self.path);wb.close()
        with self.assertRaisesRegex(ValueError,'regular MLS'):read_speech_facts(self.path)
        with self.assertRaisesRegex(ValueError,'workbook changed'):generate_speech(self.path,self.facts,options(),self.folder,False)
        self.assertEqual([],list(self.folder.glob('*PC_Speech*')))

    def test_pdf_failure_keeps_word_and_records_reason(self):
        with patch('core.office_convert.office_to_pdf',side_effect=RuntimeError('Office unavailable')):
            result=generate_speech(self.path,self.facts,options(),self.folder)
        self.assertTrue(Path(result['docx']).is_file());self.assertIsNone(result['pdf'])
        self.assertTrue(any('Office unavailable' in n for n in result['notes']))

    def test_button_preview_save_and_preference_merge_in_real_ui(self):
        from wizards.mls.wizard import MailoutWizard
        from wizards.mls.pc_speech_ui import PCSpeech
        root=tk.Tk();root.withdraw();self.addCleanup(root.destroy)
        with patch.object(session,'load',return_value={}),patch.object(session,'SETTINGS_PATH',self.folder/'session.json'), \
             patch('core.workbooks.WorkbookReadCache.resolve',return_value=str(self.path)), \
             patch('tkinter.messagebox.showerror',side_effect=AssertionError),patch('tkinter.messagebox.showinfo'), \
             patch('core.office_convert.office_to_pdf',side_effect=RuntimeError('Synthetic test has no Office export')):
            owner=MailoutWizard(root);owner.vars['excel'].set(str(self.path));owner.vars['output_folder'].set(str(self.folder))
            def descendants(w):
                for c in w.winfo_children():yield c;yield from descendants(c)
            buttons=lambda:[w for w in descendants(owner) if isinstance(w,ttk.Button) and str(w.cget('text'))=='PC Speech']
            self.assertFalse(buttons())
            if hasattr(owner,'_workflow'):owner._workflow.show('documents')
            else:owner._show_step(3)
            self.assertTrue(buttons());buttons()[0].invoke()
            window=owner._pc_speech
            def wait(predicate):
                end=time.monotonic()+15
                while not predicate() and time.monotonic()<end:root.update();time.sleep(.02)
                self.assertTrue(predicate())
            wait(lambda:window.facts is not None)
            root.deiconify();window.deiconify();window.geometry('920x700');root.update()
            scroller=window.input_scroller
            self.assertTrue(scroller._canvas.winfo_height()>=scroller.interior.winfo_reqheight()
                            or scroller._vbar.winfo_manager()=='grid')
            self.assertLessEqual(int(window.details_label.cget('wraplength')),window.details_label.master.winfo_width())
            self.assertIn('Crystalwood',window.preview.get('1.0','end'))
            for k,var in window.fields.items():var.set(options()[k])
            window.notice.set(True);window.update_preview()
            self.assertIn('Taylor Planner',window.preview.get('1.0','end'))
            save_preferences(window.source,window.output,window.facts['workbook_hash'],{'agenda':'VII','protested_rows':[18]})
            window.save();wait(lambda:hasattr(window,'last_result'))
            self.assertTrue(Path(window.last_result['docx']).is_file())
            prefs=load_preferences(window.source,window.output,window.facts['workbook_hash'])
            self.assertEqual('VII',prefs['agenda']);self.assertEqual([18],prefs['protested_rows']);self.assertEqual('12',prefs['speech']['agenda'])
            window.close();owner.destroy()

def main():
    import sys,json
    if '--office' in sys.argv:
        folder=Path(sys.argv[sys.argv.index('--office')+1]);folder.mkdir(parents=True,exist_ok=True)
        source,facts,_=case_fixture(folder);before=source.read_bytes()
        result=generate_speech(source,facts,options(),folder)
        assert result['pdf'] and not result['notes'],result
        from pypdf import PdfReader
        reader=PdfReader(result['pdf']);text=' '.join(' '.join(p.extract_text() for p in reader.pages).split())
        for expected in ('Taylor Planner','5,000','75%','35%','25%','STOP HERE'):
            assert expected in text,(expected,text)
        assert len(reader.pages)==1 and source.read_bytes()==before
        (folder/'office_verified.json').write_text(json.dumps(result,indent=2))
        print('Speech Word/PDF export verified:',result['folder'])
        return 0
    result=unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(PCSpeechTests))
    return 0 if result.wasSuccessful() else 1

if __name__=='__main__':raise SystemExit(main())
