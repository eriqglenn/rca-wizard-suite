"""Recorded plat search, source-PDF preservation, browser capture, and wizard checks."""
from __future__ import annotations
import argparse
from dataclasses import asdict
from http.server import BaseHTTPRequestHandler,ThreadingHTTPServer
import json
from pathlib import Path
import queue
import tempfile
import threading
import time
import unittest
from unittest.mock import patch

import fitz
from openpyxl import Workbook
from core import plat_search as PS

OUT=None
UI=BROWSER=LIVE=False


def fixture_pdf(text='RECORDED PLAT TEST',pages=2):
    with fitz.open() as doc:
        for n in range(pages):
            p=doc.new_page(width=792,height=612)
            p.insert_text((40,40),f'{text} - PAGE {n+1}')
            p.draw_rect(fitz.Rect(80,90,650,490),width=2)
            p.insert_text((100,510),'TEST COPY WATERMARK')
            if n%2:p.set_rotation(90)
        return doc.tobytes()


def record(name='EASTWOOD',file='114728'):
    return {'name':name,'file_number':file,'file_date':'02/13/1913','record_type':'Map','map_pages':2,'volume_page':'004042',
            'viewer_url':'https://www.cclerk.hctx.net/Applications/WebSearch/EComm/ViewEdocs.aspx?ID=public-test-id'}


def result_html(total=1):
    return f'''<html><body>{total} Record(s) Found.<table><tr>
    <td><span id="ctl_ListView1_ctrl0_lblFileNo">114728</span></td>
    <td><span id="ctl_ListView1_ctrl0_lblFileDate">02/13/1913</span></td>
    <td><span id="ctl_ListView1_ctrl0_lblIType">Map</span></td>
    <td><span id="ctl_ListView1_ctrl0_lblDesc">EASTWOOD</span><table><tr><td>Abstract</td></tr></table></td>
    <td><span id="ctl_ListView1_ctrl0_lblMapPgs">2</span></td>
    <td><a id="ctl_ListView1_ctrl0_HyperLinkVolNoEC" href="EComm/ViewEdocs.aspx?ID=map">004042</a></td>
    <td><a id="ctl_ListView1_ctrl0_HyperLinkFCEC" href="EComm/ViewEdocs.aspx?ID=separate-instrument">RP123</a></td>
    </tr></table></body></html>'''.encode()


class PlatFinderTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp=tempfile.TemporaryDirectory()
        cls.folder=Path(OUT or cls.temp.name).resolve();cls.folder.mkdir(parents=True,exist_ok=True)
        cls.pdf=fixture_pdf();cls.results=[]

    @classmethod
    def tearDownClass(cls):
        (cls.folder/'results.json').write_text(json.dumps(cls.results,indent=2),encoding='utf-8')
        cls.temp.cleanup()

    def test_search_parser_record_identity_pagination_and_no_instrument_substitution(self):
        result=PS.parse_results(result_html(100),PS.SEARCH_URL)
        self.assertEqual(result['total'],100);self.assertTrue(result['truncated'])
        self.assertEqual(len(result['records']),1)
        r=result['records'][0]
        self.assertEqual(r['name'],'EASTWOOD');self.assertEqual(r['map_pages'],2)
        self.assertTrue(r['viewer_url'].endswith('ID=map'))
        self.assertEqual(PS.parse_results(b'<html>0 Record(s) Found.</html>',PS.SEARCH_URL)['records'],[])
        with self.assertRaises(PS.PlatError):PS.parse_results(b'<html>Please log in</html>',PS.SEARCH_URL)
        with self.assertRaises(PS.PlatError):PS.parse_results(result_html().replace(b'EComm/ViewEdocs.aspx?ID=map',b'https://example.com/not-a-plat'),PS.SEARCH_URL)

    def test_search_uses_actual_form_and_validates_inputs(self):
        form=b'<form><input type="hidden" name="__VIEWSTATE" value="state"><input name="ctl00$ContentPlaceHolder1$btnSearch"><input name="actualPageInput" aria-label="Page"></form>'
        client=PS.PlatSearchClient()
        with patch.object(client,'request',side_effect=[(form,PS.SEARCH_URL),(result_html(),PS.SEARCH_URL)]) as req:
            result=client.search(volume='4',page='42')
            sent=req.call_args_list[-1].args[1]
            self.assertEqual(sent['__VIEWSTATE'],'state');self.assertEqual(sent['actualPageInput'],'42')
            self.assertEqual(sent['ctl00$ContentPlaceHolder1$txtVolNo'],'4');self.assertEqual(len(result['records']),1)
        for params in ({},{'volume':'4'},{'volume':'x','page':'42'}):
            with self.assertRaises(PS.PlatError):client.search(**params)
        for url in ('http://www.cclerk.hctx.net/test','https://example.com/plat.pdf','javascript:alert(1)'):
            with self.assertRaises(PS.PlatError):PS.clerk_url(url)
        stopped=threading.Event();stopped.set()
        with self.assertRaisesRegex(PS.PlatError,'cancelled'):PS.PlatSearchClient(stopped).request(PS.SEARCH_URL)

    def test_all_workbook_types_suggest_names_without_a_completed_calculation(self):
        for key in ('mls','mbl','mls_renewal','mbl_renewal'):
            path=self.folder/(key+'.xlsx');wb=Workbook();s=wb.active;s.title='Step 1 - Input Properties Info'
            s.append(['ALT_TAX_ID','Legal_Dscr_2']);s.append(['0611070000020','EASTWOOD 3RD PAR R/P']);s.append(['0611070000001','EASTWOOD'])
            s=wb.create_sheet('Step 2 - Main Calculation Sheet');s['E4']='#NUM!'
            ci=wb.create_sheet('Case_Info');ci.append(['File_Number','91 REN'])
            wb.save(path);wb.close();before=path.read_bytes()
            result=PS.read_plat_context(path)
            self.assertEqual(len(result['names']),2);self.assertEqual(result['reference'],'91 REN')
            self.assertEqual(result['names'][0]['search_name'],'EASTWOOD');self.assertEqual(path.read_bytes(),before)
        self.assertEqual(PS.search_name('EAST SUNNYSIDE COURT SEC 3'),'EAST SUNNYSIDE COURT')
        self.assertEqual(PS.search_name('EAST NORHILL'),'EAST NORHILL')
        self.assertTrue(PS.read_plat_context('')['warnings'])

    def test_single_pdf_exact_combined_pages_rotation_watermarks_and_unique_files(self):
        first={'record':record(),'pdf':self.pdf,'pages':2}
        second={'record':record('SECOND PLAT','2'),'pdf':fixture_pdf('SECOND',1),'pages':1}
        a=PS.save_plats([first],self.folder/'saved');b=PS.save_plats([first],self.folder/'saved')
        self.assertNotEqual(a,b);self.assertEqual(a.read_bytes(),self.pdf)
        combined=PS.save_plats([first,second],self.folder/'saved')
        with fitz.open(combined) as doc:
            self.assertEqual(len(doc),3);self.assertEqual(doc[1].rotation,90)
            self.assertIn('SECOND',doc[2].get_text())
            for p in doc:self.assertIn('TEST COPY WATERMARK',p.get_text())
        with self.assertRaises(PS.PlatError):PS.save_plats([],self.folder/'saved')
        with self.assertRaises(PS.PlatError):PS.validated_pdf(b'<html>Login required</html>')

    def test_public_county_search_for_three_neighborhoods(self):
        if not LIVE:self.skipTest('Use --live for the County public search')
        for name in ('EASTWOOD','EAST NORHILL','EAST SUNNYSIDE COURT'):
            result=PS.PlatSearchClient().search(name=name)
            self.assertTrue(result['records'],name)
            self.assertTrue(any(name in r['name'] for r in result['records']))
            (self.folder/(name.replace(' ','_')+'.json')).write_text(json.dumps(result,indent=2),encoding='utf-8')
            self.results.append({'live_search':name,'records':len(result['records'])})
        exact=PS.PlatSearchClient().search(file_number='114728')
        self.assertTrue(any(r['file_number']=='114728' for r in exact['records']))
        volume=PS.PlatSearchClient().search(volume='4',page='42')
        self.assertTrue(any(r['file_number']=='114728' for r in volume['records']))

    def test_browser_captures_actual_pdf_without_automating_login(self):
        if not BROWSER:self.skipTest('Use --browser for Chrome PDF capture against a local server')
        from core import plat_browser as PB
        content=self.pdf
        class Handler(BaseHTTPRequestHandler):
            def log_message(self,*args):pass
            def do_GET(self):
                if self.path=='/plat.pdf':
                    body=content;kind='application/pdf'
                else:body=b'<html><iframe src="/plat.pdf"></iframe></html>';kind='text/html'
                self.send_response(200);self.send_header('Content-Type',kind);self.send_header('Content-Length',str(len(body)));self.end_headers();self.wfile.write(body)
        server=ThreadingHTTPServer(('127.0.0.1',0),Handler)
        threading.Thread(target=server.serve_forever,daemon=True).start()
        events=queue.Queue()
        browser=PB.PlatBrowser(events,profile=self.folder/('browser_'+str(time.monotonic_ns())),headless=True)
        try:
            with patch.object(PB,'CLERK_HOSTS',{'127.0.0.1'}),patch.object(PB,'clerk_url',side_effect=lambda u:u):
                r=dict(record(),viewer_url=f'http://127.0.0.1:{server.server_port}/viewer')
                browser.start(r);deadline=time.monotonic()+35;got=None;errors=[]
                while time.monotonic()<deadline:
                    try:kind,value=events.get(timeout=.2)
                    except queue.Empty:continue
                    if kind=='downloaded':got=value
                    if kind=='browser_error':errors.append(value)
                    if kind=='browser_done':break
                browser.close();browser.thread.join(timeout=10)
                self.assertIsNotNone(got,errors)
                self.assertEqual(got['pdf'],content);self.assertEqual(got['pages'],2)
                self.results.append({'browser_pdf_capture':True,'county_login_tested':False})
        finally:
            browser.close();server.shutdown();server.server_close()

    def test_all_four_wizards_search_import_preview_export_attach_and_session(self):
        if not UI:self.skipTest('Use --ui for wizard controls')
        import tkinter as tk
        import importlib
        from core import plat_search_ui as PU,session
        names={'mls':'MailoutWizard','mbl':'MBLWizard','mls_renewal':'MLSRenewalWizard','mbl_renewal':'MBLRenewalWizard'}
        root=tk.Tk();root.withdraw()
        def descendants(widget):
            for child in widget.winfo_children():yield child;yield from descendants(child)
        def wait(predicate):
            deadline=time.monotonic()+12
            while not predicate() and time.monotonic()<deadline:root.update();time.sleep(.02)
            self.assertTrue(predicate())
        def fail(title,message,**kwargs):raise AssertionError(f'{title}: {message}')
        imported=self.folder/'manual_plat.pdf';imported.write_bytes(fixture_pdf('MANUAL',1))
        def fake_start(browser,rec):
            browser.events.put(('downloaded',{'record':rec,'pdf':self.pdf,'pages':2,'source_url':rec['viewer_url']}))
            browser.events.put(('browser_done',None))
        try:
            with patch.object(session,'SETTINGS_PATH',self.folder/'ui-session.json'),patch.object(session,'load',return_value={}),patch('tkinter.messagebox.showerror',fail),\
                 patch.object(PU,'read_plat_context',return_value={'names':[{'name':'EASTWOOD','search_name':'EASTWOOD','properties':2}],'warnings':[]}),\
                 patch.object(PU.PlatSearchClient,'search',return_value={'records':[record()],'total':1,'truncated':False}),patch.object(PU.PlatBrowser,'start',fake_start):
                for key,cls in names.items():
                    owner=None
                    try:
                        owner=getattr(importlib.import_module(f'wizards.{key}.wizard'),cls)(root)
                        folder=self.folder/'ui'/key;folder.mkdir(parents=True,exist_ok=True)
                        if hasattr(owner,'vars'):
                            owner.vars['output_folder'].set(str(folder));owner.vars['map'].set('ORIGINAL_MAP.pdf');owner.vars['tax_map'].set('ORIGINAL_TAX.pdf')
                        else:owner.files.update(out=str(folder),map='ORIGINAL_MAP.pdf',tax_map='ORIGINAL_TAX.pdf')
                        buttons=[w for w in descendants(owner) if 'text' in w.keys() and w.cget('text')=='Find Subdivision Plats']
                        self.assertEqual(len(buttons),1);buttons[0].invoke();finder=owner._plat_finder
                        wait(lambda:len(finder.records)==1)
                        finder.results.selection_set('0');finder.selected_record();finder.download_button.invoke()
                        wait(lambda:len(finder.items)==1)
                        finder.turn_page(1);self.assertEqual(finder.page,1)
                        with patch('tkinter.filedialog.askopenfilenames',return_value=[str(imported)]):finder.add_files()
                        self.assertEqual(len(finder.items),2)
                        finder.add_item(dict(finder.items[0]));self.assertEqual(len(finder.items),2)
                        export=folder/'export.pdf'
                        with patch('tkinter.filedialog.asksaveasfilename',return_value=str(export)):finder.export_button.invoke()
                        with fitz.open(export) as doc:self.assertEqual(len(doc),3)
                        finder.use_button.invoke()
                        values=owner.vars if hasattr(owner,'vars') else owner.files
                        def value(k):return values[k].get() if hasattr(owner,'vars') else values[k]
                        selected=value('subdivision_plat');self.assertTrue(Path(selected).exists())
                        self.assertEqual(value('map'),'ORIGINAL_MAP.pdf');self.assertEqual(value('tax_map'),'ORIGINAL_TAX.pdf')
                        saved=json.loads((self.folder/'ui-session.json').read_text())
                        self.assertEqual(saved[key]['subdivision_plat'],selected)
                        with fitz.open(selected) as doc:self.assertEqual(len(doc),3)
                        self.results.append({'wizard_ui':key,'search_download_import_export_use':True})
                    finally:
                        if owner is not None:
                            finder=getattr(owner,'_plat_finder',None)
                            if finder is not None and finder.winfo_exists():finder.close()
                            owner.destroy()
                        root.update()
        finally:root.destroy()


def main():
    global OUT,UI,BROWSER,LIVE
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--out');parser.add_argument('--ui',action='store_true');parser.add_argument('--browser',action='store_true');parser.add_argument('--live',action='store_true')
    args=parser.parse_args();OUT,UI,BROWSER,LIVE=args.out,args.ui,args.browser,args.live
    result=unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(PlatFinderTests))
    return 0 if result.wasSuccessful() else 1


if __name__=='__main__':raise SystemExit(main())
