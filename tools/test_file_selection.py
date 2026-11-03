"""Exercise typed file selections through the actual wizard widgets."""
import importlib
import json
import os
from pathlib import Path
import shutil
import sys
import tempfile
import unittest
from unittest.mock import patch

WF=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(WF))
from core import session, office_convert, errorlog
from core.embedding import SUITE_VERSION
from core.paths import sample_input_dir
from core.workbooks import CALCULATION_TEMPLATES

WIZARDS={'mls':'MailoutWizard','mbl':'MBLWizard','mls_renewal':'MLSRenewalWizard','mbl_renewal':'MBLRenewalWizard'}

def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)

class FileSelectionTests(unittest.TestCase):
    def setUp(self):
        import app
        self.temp=tempfile.TemporaryDirectory(prefix='rca_ui_paths_')
        self.addCleanup(self.temp.cleanup)
        self.folder=Path(self.temp.name)
        for p in (patch.object(Path,'home',return_value=self.folder/'profile'),
                  patch.object(session,'SETTINGS_PATH',self.folder/'settings.json'),
                  patch.object(app.SuiteApp,'_setup_summary',return_value='Test',create=True)):
            p.start()
            self.addCleanup(p.stop)
        self.root=app.SuiteApp()
        self.root.withdraw()
        self.addCleanup(self.root.destroy)

    def open(self,key):
        self.root.open_wizard(key)
        self.root.update()
        return next(w for w in descendants(self.root) if type(w).__name__==WIZARDS[key])

    def entry(self,wizard,label):
        for w in descendants(wizard):
            if w.winfo_class() in ('Label','TLabel') and str(w.cget('text'))==label:
                row=w.grid_info().get('row')
                entries=[x for x in w.master.grid_slaves(row=row,column=1) if x.winfo_class() in ('Entry','TEntry')]
                if len(entries)==1:return entries[0]
        raise AssertionError('Missing entry: '+label)

    def type(self,wizard,label,value):
        e=self.entry(wizard,label)
        e.delete(0,'end')
        e.insert(0,str(value))
        self.assertEqual(e.get(),str(value))

    def value(self,wizard,key):
        return wizard.vars[key].get() if hasattr(wizard,'vars') else wizard.files[key]

    def page(self,wizard,key):
        (wizard._step_files if key=='mls' else wizard._select_files_page)()
        self.root.update()

    def test_typed_and_cleared_paths_survive_navigation_in_all_four_wizards(self):
        for key in WIZARDS:
            with self.subTest(wizard=key):
                w=self.open(key)
                field='output_folder' if key=='mls' else 'out'
                chosen=str(self.folder/key/'typed output')
                self.type(w,'Output folder',chosen)
                self.assertEqual(self.value(w,field),chosen)
                self.type(w,'Tax map (required)',self.folder/'tax.pdf')
                self.type(w,'Tax map (required)','')
                self.assertEqual(self.value(w,'tax_map'),'')
                (w._step_generate if key=='mls' else w._generate_page)()
                self.page(w,key)
                self.assertEqual(self.entry(w,'Output folder').get(),chosen)
                self.assertEqual(self.entry(w,'Tax map (required)').get(),'')
                self.root.go_home()

    def test_typed_workbook_switches_the_workbook_used_for_generation(self):
        for key in WIZARDS:
            with self.subTest(wizard=key):
                w=self.open(key)
                source=sample_input_dir(key)/(Path(CALCULATION_TEMPLATES[key]).stem+'_FILLED_SAMPLE.xlsx')
                first=self.folder/(key+'_first.xlsx')
                second=self.folder/(key+'_second.xlsx')
                shutil.copy2(source,first)
                shutil.copy2(source,second)
                label='Calculation workbook' if key=='mls' else ('MBL calculation workbook' if key=='mbl' else 'Renewal calculation workbook')
                self.type(w,label,first)
                self.assertEqual(Path(w._wb()).resolve(),first.resolve())
                self.type(w,label,second)
                self.assertEqual(Path(w._wb()).resolve(),second.resolve())
                self.root.go_home()

    def test_browse_after_typing_and_typing_after_browse_use_latest_selection(self):
        for key in ('mbl','mls_renewal','mbl_renewal'):
            with self.subTest(wizard=key):
                w=self.open(key)
                self.type(w,'Output folder',self.folder/'typed first')
                e=self.entry(w,'Output folder')
                button=e.master.grid_slaves(row=e.grid_info()['row'],column=2)[0]
                mod=importlib.import_module('wizards.'+key+'.wizard')
                with patch.object(mod.filedialog,'askdirectory',return_value=str(self.folder/'browsed')):
                    button.invoke()
                self.assertEqual(w.files['out'],str(self.folder/'browsed'))
                self.assertEqual(e.get(),w.files['out'])
                self.type(w,'Output folder',self.folder/'typed last')
                w._save_session()
                self.assertEqual(session.load(key)['out'],str(self.folder/'typed last'))
                self.root.go_home()

    def test_renewal_packet_uses_typed_inputs_and_output(self):
        from pypdf import PdfReader
        import fitz
        key='mbl_renewal'
        w=self.open(key)
        mod=importlib.import_module('wizards.'+key+'.wizard')
        source=sample_input_dir(key)/(Path(CALCULATION_TEMPLATES[key]).stem+'_FILLED_SAMPLE.xlsx')
        workbook=self.folder/'typed workbook.xlsx'
        shutil.copy2(source,workbook)
        self.type(w,'Renewal calculation workbook',workbook)
        for label,name in [('Boundary map','boundary'),('Tax map (required)','tax'),('Original ordinance (required)','original')]:
            pdf=self.folder/(name+'.pdf')
            with fitz.open() as d:
                d.new_page().insert_text((72,72),'Typed attachment: '+name)
                d.save(pdf)
            self.type(w,label,pdf)
        destination=self.folder/'typed final output'
        self.type(w,'Output folder',destination)
        def export(src,dest):
            result=Path(dest)/(Path(src).stem+'.pdf')
            with fitz.open() as d:
                d.new_page().insert_text((72,72),'Office fixture: '+Path(src).name)
                d.save(result)
            return result
        with patch.object(office_convert,'office_to_pdf',side_effect=export), \
             patch.object(mod.messagebox,'showerror') as failed, \
             patch.object(mod.messagebox,'showinfo'),patch.object(mod.os,'startfile'):
            w._generate_page()
            button=next(x for x in descendants(w) if x.winfo_class()=='Button' and x.cget('text')=='Generate MBL Renewal RCA Packet PDF')
            button.invoke()
            failed.assert_not_called()
        output=destination/'MBL_Renewal_RCA_Packet.pdf'
        self.assertEqual(len(PdfReader(output).pages),5)
        content=' '.join(p.extract_text() or '' for p in PdfReader(output).pages)
        for token in ('boundary','tax','original'):self.assertIn('Typed attachment: '+token,content)
        self.assertEqual(session.load(key)['out'],str(destination))

    def test_mls_packet_failure_reaches_delayed_error_callback(self):
        w=self.open('mls')
        mod=importlib.import_module('wizards.mls.wizard')
        callbacks=[]
        class ImmediateThread:
            def __init__(self,target,**kw):self.target=target
            def start(self):self.target()
        error=RuntimeError('The actual Office export failed')
        with patch.object(w,'_validate_paths'),patch.object(w,'_wb',return_value='workbook.xlsx'), \
             patch.object(mod,'generate_rca_packet_pdf',side_effect=error), \
             patch.object(mod.threading,'Thread',ImmediateThread), \
             patch.object(w,'after',side_effect=lambda delay,callback:callbacks.append(callback)), \
             patch.object(w,'_error') as show_error:
            w._run_packet()
            self.assertEqual(len(callbacks),1)
            callbacks[0]()  # run AFTER the exception handler has exited
            show_error.assert_called_once_with(error)

    def test_cleared_output_folder_is_rejected_before_generation(self):
        for key in ('mbl','mls_renewal','mbl_renewal'):
            with self.subTest(wizard=key):
                w=self.open(key)
                mod=importlib.import_module('wizards.'+key+'.wizard')
                self.type(w,'Output folder','')
                with patch.object(w,'_wb',return_value='workbook.xlsx'), \
                     patch.object(mod,'generate_rca_packet_pdf') as generate, \
                     patch.object(w,'_notify_done'),patch.object(w,'_fail') as failure:
                    w._rca_packet()
                    generate.assert_not_called()
                    failure.assert_called_once()
                    self.assertIn('output folder',str(failure.call_args.args[0]).lower())
                self.root.go_home()

    def test_error_log_identifies_the_actual_build_and_launch_mode(self):
        for frozen in (False,True):
            with self.subTest(frozen=frozen),patch.object(sys,'frozen',frozen,create=True):
                out=self.folder/str(frozen)
                log=errorlog.log_exception('test wizard',out,ValueError('Original error'))
                text=log.read_text(encoding='utf-8')
                self.assertIn(SUITE_VERSION,text)
                self.assertIn('packaged EXE' if frozen else 'Python source launcher',text)
                self.assertIn(sys.executable,text)
                self.assertIn('ValueError: Original error',text)

if __name__=='__main__':unittest.main(verbosity=2)
