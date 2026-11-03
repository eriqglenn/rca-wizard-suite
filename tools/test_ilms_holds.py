"""ILMS hold tests use local records and an intercepted browser fixture, never live writes."""
import copy
import importlib
import json
import queue
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch
from openpyxl import Workbook
from core.ilms_holds import *
from core.ilms_browser import ILMSPage,ILMS_URL,ILMSError,ILMSBrowser,browser_error_message

def workbook(path):
    b=Workbook();s=b.active;s.title='Step 1 - Input Properties Info'
    s.append(['TAX_ID','Site_addr_1','Legal_Dscr_1','Land Use','str_unit'])
    for row in [('0401047810017','12107 CRYSTALWOOD DR','LT 1','SFR',''),
                ('0401047810017','12107 CRYSTALWOOD DRIVE','LT 2','SFR',''),
                ('0401047810018','12109 CRYSTALWOOD DR','LT 3','EXC',''),
                ('0401047810019','0 CRYSTALWOOD DR','LT 4','VAC',''),
                ('0401047810020','12111 CRYSTALWOOD DR','LT 5','SFR','1')]:s.append(row)
    b.create_sheet('Step 2 - Calculation Sheet')['B2']='121';b.save(path);b.close();return path

class Adapter:
    def __init__(self,address='12107 CRYSTALWOOD DR',rows=None):
        self.address=address;self.rows=rows or [['01/01/2007','PLANNING','HARD HOLD','','GE','old hold','OLD']]
        self.calls=0;self.reason='';self.fail=False
    def find(self,address):return {'address':self.address}
    def holds(self):return copy.deepcopy(self.rows)
    matching_count=staticmethod(ILMSPage.matching_count)
    def prepare(self,reason):self.reason=reason
    def assert_address(self,address):
        if not same_address(self.address,address):raise ValueError('Wrong situs')
    def submit(self):
        self.calls+=1;self.rows.append(['09/24/2026','PLANNING','HARD HOLD','','GE',self.reason,'TEST'])
    def verify_added(self,reason,before):
        if self.fail:raise TimeoutError('Connection lost after Add')
        assert self.matching_count(self.rows,reason)==before+1
        return self.holds()

class ILMSTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup);self.folder=Path(self.tmp.name)
        self.property={'id':'p','key':address_key('12107 CRYSTALWOOD DR'),'address':'12107 CRYSTALWOOD DR','selected':True,'rows':[2,3],'status':'Address matched'}
        self.reason=hold_reason('121')
    def journal(self):return Journal(self.folder,'mls','121')

    def test_reason_and_application_validation(self):
        self.assertEqual('MLS/MBL 121 call planning 6637',self.reason)
        self.assertEqual('MLS/MBL 187REN call planning 6637',hold_reason('MLS 187 REN'))
        for value in ('','../121','unknown','=1+1'):
            with self.assertRaises(ValueError):hold_reason(value)

    def batch_plan(self):
        return {'application':'121','properties':[dict(self.property,address=f'{n} CRYSTALWOOD DR',key=address_key(f'{n} CRYSTALWOOD DR'),status='Needs review') for n in (12107,12109,12111)]}

    def test_batch_navigation_failure_pauses_and_preserves_completed_holds(self):
        events=queue.Queue();worker=ILMSBrowser(events);plan=self.batch_plan();journal=self.journal()
        place_one(Adapter(),plan['properties'][0],self.reason,journal)
        before=copy.deepcopy(journal.entry(plan['properties'][0]['key']))
        adapter=Adapter()
        with patch.object(adapter,'find',side_effect=ILMSError('Could not identify 126 Situs Holds')) as find:
            worker.process_batch(adapter,'automatic',plan,journal)
        self.assertEqual(1,find.call_count);self.assertEqual(0,adapter.calls)
        self.assertEqual(['Completed in this batch','Paused','Not attempted'],[p['status'] for p in plan['properties']])
        self.assertEqual(before,journal.entry(plan['properties'][0]['key']))
        self.assertIn('Paused.',list(events.queue)[-1][1]['message'])
        # Resume uses the journal, not stale UI status, and visits only pending addresses.
        adapters={p['address']:Adapter(p['address']) for p in plan['properties']}
        from core.ilms_holds import place_one as real_place
        with patch('core.ilms_browser.place_one',side_effect=lambda a,p,r,j,s:real_place(adapters[p['address']],p,r,j,s)):
            worker.process_batch(adapter,'automatic',plan,journal)
        self.assertEqual([0,1,1],[a.calls for a in adapters.values()])
        self.assertEqual(['Completed in this batch','Hold added','Hold added'],[p['status'] for p in plan['properties']])

    def test_batch_uncertain_add_is_reconciled_before_later_properties(self):
        worker=ILMSBrowser(queue.Queue());plan=self.batch_plan();journal=self.journal()
        adapters={p['address']:Adapter(p['address']) for p in plan['properties']}
        first=next(iter(adapters.values()));first.fail=True
        from core.ilms_holds import place_one as real_place
        with patch('core.ilms_browser.place_one',side_effect=lambda a,p,r,j,s:real_place(adapters[p['address']],p,r,j,s)):
            worker.process_batch(first,'automatic',plan,journal)
            self.assertEqual(['Needs review','Not attempted','Not attempted'],[p['status'] for p in plan['properties']])
            self.assertEqual([1,0,0],[a.calls for a in adapters.values()])
            worker.process_batch(first,'automatic',plan,journal)
        self.assertEqual([1,1,1],[a.calls for a in adapters.values()])
        self.assertEqual('Confirmed after interruption',plan['properties'][0]['status'])

    def test_stop_before_batch_does_not_mark_other_properties_as_failures(self):
        worker=ILMSBrowser(queue.Queue());worker.stop.set();plan=self.batch_plan();adapter=Adapter()
        worker.process_batch(adapter,'automatic',plan,self.journal())
        self.assertEqual(['Paused','Not attempted','Not attempted'],[p['status'] for p in plan['properties']])
        self.assertEqual(0,adapter.calls)

    def test_browser_disconnect_is_actionable_and_restart_discards_stale_commands(self):
        self.assertIn('Start / Resume',browser_error_message(RuntimeError('Page.wait_for_timeout: Target page, context or browser has been closed')))
        worker=ILMSBrowser(queue.Queue());worker.commands.put(('automatic',{'obsolete':True}))
        with patch('core.ilms_browser.threading.Thread') as thread:
            worker.start()
        self.assertTrue(worker.commands.empty());thread.return_value.start.assert_called_once()

    def test_duplicate_and_delayed_situs_navigation(self):
        from playwright.sync_api import sync_playwright
        variants=[
            "const originalSelected=selected;selected=function(n){originalSelected(n);identity.insertAdjacentHTML('beforeend','<button onclick=\"holds()\">126 Situs Holds</button>')};",
            "const originalSelected=selected;selected=function(n){originalSelected(n);identity.querySelector('button').remove();setTimeout(()=>identity.insertAdjacentHTML('beforeend','<button onclick=\"holds()\">126 Situs Holds</button>'),700)};",
            PROJECT_EMPTY_FIXTURE.replace('role=\"treeitem\"','role=\"button\"')
        ]
        with sync_playwright() as pw:
            browser=pw.chromium.launch(channel='chrome',headless=True)
            try:
                for index,variant in enumerate(variants):
                    with self.subTest(variant=index):
                        page=browser.new_page(viewport={'width':1500,'height':1000})
                        fixture=FIXTURE.replace('search();</script>',variant+'search();</script>')
                        page.route('**/*',lambda route:route.fulfill(status=200,content_type='text/html',body=fixture))
                        page.goto(ILMS_URL);adapter=ILMSPage(page)
                        journal=Journal(self.folder/str(index),'mls','121')
                        self.assertEqual('Hold added',place_one(adapter,self.property,self.reason,journal)['status'])
                        self.assertEqual(1,page.evaluate('window.addClicks'));page.close()
            finally:browser.close()

    def test_address_normalization_preserves_direction_unit_fraction(self):
        self.assertTrue(same_address('12107 Crystalwood Drive','12107 CRYSTALWOOD DR 77013'))
        self.assertTrue(same_address('1700 West Main Street','1700 W MAIN ST'))
        self.assertFalse(same_address('1700 W MAIN ST','1700 MAIN ST'))
        self.assertFalse(same_address('100 TEST ST UNIT 1','100 TEST ST UNIT 2'))
        self.assertFalse(same_address('100 1/2 TEST ST','100 TEST ST'))
        self.assertEqual('E',address_parts('100 TEST ST E UNIT 1')['suffix'])
        self.assertEqual('',address_key('0 TEST ST'));self.assertEqual('',address_key('100-200 TEST ST'))

    def test_all_four_loaders_include_exc_group_lots_and_flag_zero(self):
        path=workbook(self.folder/'case.xlsx');before=path.read_bytes()
        for key in WIZARDS:
            plan=load_plan(path,key);self.assertEqual(4,len(plan['properties']))
            self.assertEqual([2,3],plan['properties'][0]['rows'])
            self.assertTrue(plan['properties'][1]['selected']);self.assertEqual('EXC',plan['properties'][1]['land_use'])
            self.assertFalse(plan['properties'][2]['selected']);self.assertEqual('Needs address',plan['properties'][2]['status'])
            self.assertTrue(plan['properties'][3]['address'].endswith('UNIT 1'))
        self.assertEqual(before,path.read_bytes())

    def test_add_despite_existing_identical_reason_and_preserve_old_holds(self):
        adapter=Adapter(rows=[['01/01/2007','PLANNING','HARD HOLD','','GE',self.reason,'OLD']]);before=adapter.holds()
        result=place_one(adapter,self.property,self.reason,self.journal())
        self.assertEqual('Hold added',result['status']);self.assertEqual(1,adapter.calls)
        self.assertEqual(before,adapter.rows[:1]);self.assertEqual(2,adapter.matching_count(adapter.rows,self.reason))

    def test_restart_does_not_repeat_a_completed_add(self):
        adapter=Adapter();place_one(adapter,self.property,self.reason,self.journal())
        result=place_one(adapter,self.property,self.reason,self.journal())
        self.assertEqual(1,adapter.calls);self.assertEqual('Completed in this batch',result['status'])

    def test_connection_lost_after_add_is_reconciled_without_another_click(self):
        adapter=Adapter();adapter.fail=True
        with self.assertRaises(ValueError):place_one(adapter,self.property,self.reason,self.journal())
        self.assertEqual('uncertain',self.journal().entry(self.property['key'])['state'])
        result=place_one(adapter,self.property,self.reason,self.journal())
        self.assertEqual('Confirmed after interruption',result['status']);self.assertEqual(1,adapter.calls)

    def test_uncertain_no_confirmation_is_not_resubmitted(self):
        j=self.journal();j.record(self.property['key'],state='submitting',before_count=0)
        adapter=Adapter()
        with self.assertRaisesRegex(ValueError,'uncertain'):place_one(adapter,self.property,self.reason,j)
        self.assertEqual(0,adapter.calls)

    def test_mismatched_situs_unselected_and_cancelled_never_submit(self):
        a=Adapter(address='12109 CRYSTALWOOD DR')
        with self.assertRaises(ValueError):place_one(a,self.property,self.reason,self.journal())
        self.assertEqual(0,a.calls)
        p=copy.deepcopy(self.property);p['selected']=False;self.assertEqual('Not selected',place_one(a,p,self.reason,self.journal())['status'])
        stop=threading.Event();stop.set()
        with self.assertRaises(InterruptedError):place_one(a,self.property,self.reason,self.journal(),stop)
        self.assertEqual(0,a.calls)

    def test_journal_lock_and_case_specific_report(self):
        j=self.journal();j.acquire()
        try:
            with self.assertRaises(ValueError):self.journal().acquire()
            path=j.report([self.property]);self.assertIn('12107 CRYSTALWOOD DR',path.read_text(encoding='utf-8-sig'))
        finally:j.close()
        again=self.journal();again.acquire();again.close()

    def test_all_four_import_row_buttons_and_selection_controls(self):
        import tkinter as tk
        from core import session
        root=tk.Tk();root.withdraw();self.addCleanup(root.destroy)
        classes={'mls':'MailoutWizard','mbl':'MBLWizard','mls_renewal':'MLSRenewalWizard','mbl_renewal':'MBLRenewalWizard'}
        with patch.object(session,'load',return_value={}),patch.object(session,'save'),patch('core.ilms_ui.ILMSPanel.start_automatic'):
            for key,name in classes.items():
                owner=getattr(importlib.import_module('wizards.'+key+'.wizard'),name)(root)
                owner.pack(fill='both',expand=True)
                def descendants(widget):
                    for child in widget.winfo_children():
                        yield child
                        yield from descendants(child)
                buttons=[w for w in descendants(owner) if 'text' in w.keys()]
                entry=next(w for w in buttons if w['text']=='ILMS Holds')
                importer=next(w for w in buttons if w['text']=='Import Application / Enter Boundary / Fill Calculation Sheet')
                self.assertIs(entry.master,importer.master)
                self.assertEqual('left',entry.pack_info()['side'])
                self.assertEqual([importer,entry],entry.master.pack_slaves()[:2])
                entry.invoke();dialog=owner._ilms_window
                entry.invoke();self.assertIs(dialog,owner._ilms_window)
                pane=owner._ilms_panel;pane.workbook.set(str(workbook(self.folder/(key+'.xlsx'))));pane.load();root.update()
                self.assertEqual(4,len(pane.tree.get_children()));self.assertEqual(self.reason,pane.reason.get())
                p=pane.plan['properties'][1];pane.tree.selection_set(p['id']);pane.toggle();self.assertFalse(p['selected'])
                pane.workbook.set('another.xlsx');self.assertIsNone(pane.plan)
                dialog.destroy();owner.destroy();root.update()

    def test_automatic_start_loads_and_queues_one_run_without_confirmation(self):
        import tkinter as tk
        from types import SimpleNamespace
        from unittest.mock import Mock
        from core.ilms_ui import ILMSPanel
        root=tk.Tk();root.withdraw();self.addCleanup(root.destroy)
        owner=SimpleNamespace(_boundary_map_wizard_key='mls',vars={},files={})
        panel=ILMSPanel(root,owner);panel.workbook.set(str(workbook(self.folder/'auto.xlsx')))
        panel.output.set(str(self.folder))
        with patch('core.ilms_ui.ILMSBrowser') as browser,patch('core.ilms_ui.messagebox.askyesno') as confirm:
            panel.start_automatic();panel.start_automatic()
            browser.return_value.start.assert_called_once()
            browser.return_value.send.assert_called_once()
            kind,data=browser.return_value.send.call_args.args
            self.assertEqual('automatic',kind);self.assertEqual('121',data['application'])
            self.assertEqual(3,sum(p['selected'] for p in data['properties']))
            self.assertTrue(data['workbook_sha256']);confirm.assert_not_called()
            panel.stop();browser.return_value.stop.set.assert_called_once()
        panel.destroy()

    def test_login_gate_never_treats_password_form_as_authenticated(self):
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            browser=pw.chromium.launch(channel='chrome',headless=True)
            try:
                page=browser.new_page()
                page.route('**/*',lambda route:route.fulfill(status=200,content_type='text/html',body='<nav>Address Search</nav><input type="password">'))
                page.goto(ILMS_URL);adapter=ILMSPage(page)
                self.assertFalse(adapter.signed_in())
                page.locator('input').evaluate('(e)=>e.remove()')
                self.assertTrue(adapter.signed_in());adapter.wait_for_sign_in(seconds=1)
                stop=threading.Event();stop.set();adapter.cancel=stop
                with self.assertRaises(InterruptedError):adapter.wait_for_sign_in(seconds=1)
            finally:browser.close()

    def test_project_selection_collapsed_situs_and_no_existing_holds(self):
        from playwright.sync_api import sync_playwright
        fixture=FIXTURE.replace("search();</script>", PROJECT_EMPTY_FIXTURE+"search();</script>")
        with sync_playwright() as pw:
            browser=pw.chromium.launch(channel='chrome',headless=True)
            try:
                page=browser.new_page(viewport={'width':1500,'height':1000})
                page.route('**/*',lambda route:route.fulfill(status=200,content_type='text/html',body=fixture))
                page.goto(ILMS_URL);adapter=ILMSPage(page)
                result=place_one(adapter,self.property,self.reason,self.journal())
                self.assertEqual('Hold added',result['status'])
                self.assertEqual(1,page.evaluate('window.addClicks'))
                self.assertEqual(1,len(adapter.holds()))
                self.assertEqual(1,page.evaluate('window.projectClicks'))
            finally:browser.close()

    def test_next_property_returns_to_address_search_after_verified_hold(self):
        from playwright.sync_api import sync_playwright
        fixture=FIXTURE.replace("search();</script>", MULTI_ADDRESS_FIXTURE+"search();</script>")
        with sync_playwright() as pw:
            browser=pw.chromium.launch(channel='chrome',headless=True)
            try:
                page=browser.new_page(viewport={'width':1500,'height':1000})
                page.route('**/*',lambda route:route.fulfill(status=200,content_type='text/html',body=fixture))
                page.goto(ILMS_URL);adapter=ILMSPage(page);journal=self.journal()
                for number in (12107,12109):
                    prop=dict(self.property,address=f'{number} CRYSTALWOOD DR',key=address_key(f'{number} CRYSTALWOOD DR'))
                    result=place_one(adapter,prop,self.reason,journal)
                    self.assertEqual('Hold added',result['status'])
                    self.assertEqual(1,len(adapter.holds()))
                    page.evaluate("window.sequence.push('verified:'+selectedNumber)")
                self.assertEqual(['search','search','add:12107','verified:12107','search','add:12109','verified:12109'],page.evaluate('window.sequence'))
                self.assertEqual(2,page.evaluate('window.addClicks'))
                self.assertEqual([1,1],page.evaluate('Object.values(window.byAddress).map(rows=>rows.length)'))
            finally:browser.close()

    def test_layout_table_and_custom_element_captions(self):
        from playwright.sync_api import sync_playwright
        html='''<h3>Address Search</h3><table><tr><td>Number<br><input id="number"></td><td><kendo-label>Street Name</kendo-label><input id="street"></td></tr></table>
        <div class="k-grid" role="grid"><table><tr><td>Number<input id="wrong"></td></tr></table></div>
        <h3>Enforce Additional Situs Holds</h3><table><tr><td>Permit Type</td><td><input id="permit"></td><td><input readonly id="description"></td></tr>
        <tr><td><strong>Department</strong></td><td><input id="dept"></td></tr>
        <tr><td>Hold Type</td><td><input id="holdCode"></td><td><select id="hold"><option>Hard Hold</option></select></td></tr>
        <tr><td>Reason</td><td colspan="2"><input id="reason"></td></tr></table>'''
        with sync_playwright() as pw:
            browser=pw.chromium.launch(channel='chrome',headless=True)
            try:
                page=browser.new_page();page.route('**/*',lambda route:route.fulfill(status=200,content_type='text/html',body=html))
                page.goto(ILMS_URL);adapter=ILMSPage(page)
                for caption,expected in [('Number','number'),('Street Name','street')]:
                    self.assertEqual(expected,adapter.field(caption,section='Address Search').get_attribute('id'))
                for caption,expected in [('Permit Type','permit'),('Department','dept'),('Hold Type','holdCode'),('Reason','reason')]:
                    self.assertEqual(expected,adapter.field(caption,section='Enforce Additional Situs Holds').get_attribute('id'))
                self.assertEqual('hold',adapter.field('Hold Type',kind='combo',section='Enforce Additional Situs Holds').get_attribute('id'))
                adapter.fill('Number','1651',section='Address Search')
                self.assertEqual('1651',page.locator('#number').input_value());self.assertEqual('',page.locator('#wrong').input_value())
            finally:browser.close()

    def test_retained_results_project_header_and_delayed_next_search(self):
        from playwright.sync_api import sync_playwright
        fixture=FIXTURE.replace('search();</script>',RETAINED_FIXTURE+'search();</script>')
        with sync_playwright() as pw:
            browser=pw.chromium.launch(channel='chrome',headless=True)
            try:
                page=browser.new_page(viewport={'width':1500,'height':1000})
                page.route('**/*',lambda route:route.fulfill(status=200,content_type='text/html',body=fixture))
                page.goto(ILMS_URL);adapter=ILMSPage(page);journal=self.journal()
                for number in (12107,12109):
                    prop=dict(self.property,address=f'{number} CRYSTALWOOD DR',key=address_key(f'{number} CRYSTALWOOD DR'))
                    self.assertEqual('Hold added',place_one(adapter,prop,self.reason,journal)['status'])
                self.assertEqual(2,page.evaluate('window.addClicks'))
                self.assertEqual(2,page.evaluate('window.projectClicks'))
                self.assertEqual([12107,12109],page.evaluate('window.savedAddresses'))
                self.assertEqual(['',''],page.evaluate('window.submittedZips'))
            finally:browser.close()

    def test_browser_controls_against_local_intercepted_fixture(self):
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            browser=pw.chromium.launch(channel='chrome',headless=True)
            try:
                page=browser.new_page(viewport={'width':1500,'height':1000})
                page.route('**/*',lambda route:route.fulfill(status=200,content_type='text/html',body=FIXTURE))
                page.goto(ILMS_URL);adapter=ILMSPage(page)
                match=adapter.find(self.property['address']);self.assertTrue(same_address(match['address'],self.property['address']))
                result=place_one(adapter,self.property,self.reason,self.journal())
                self.assertEqual('Hold added',result['status']);self.assertEqual(1,page.evaluate('window.addClicks'))
                self.assertEqual(3,len(adapter.holds()))
            finally:browser.close()

FIXTURE=r'''<!doctype html><html><head><style>
body{font:16px Arial} .search{display:grid;grid-template-columns:repeat(5,180px);gap:20px}
.search label{display:block}.search input,.search select{width:165px;height:25px} .line{display:flex;align-items:center;gap:12px;margin:15px}.line label{width:120px} .line input{width:100px} .line input.wide{width:320px} th,td{padding:8px} .k-grid{margin-top:20px}
</style></head><body><nav><button onclick="search()">Address Search</button></nav><div id="identity"></div><main id="main"></main><script>
window.addClicks=0;window.recordRows=[['2007','PLANNING','HARD HOLD','','GE','OLD MLS HOLD','OLD'],['2020','PLANNING','HARD HOLD','','GE','MLS/MBL 121 call planning 6637','OLD']];
const esc=s=>String(s).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('"','&quot;');
const input=(label,id)=>`<div><label>${label}</label><input id="${id}"></div>`;
const select=(label,id,options)=>`<div><label>${label}</label><select id="${id}">${options.map(o=>`<option>${o}</option>`).join('')}</select></div>`;
function search(){identity.innerHTML='';main.innerHTML='<h2>Find Situs</h2><h3>Address Search</h3><div class="search">'+input('Number','num')+select('Fraction','frac',['','1/2'])+select('Direction','pre',['','WEST'])+input('Street Name','street')+select('Street Type','typ',['','DRIVE','STREET'])+select('Direction','post',['','EAST'])+input('Unit Type','unitType')+input('Unit No.','unit')+input('Zip Code','zip')+input('Zip Plus','zipPlus')+'</div><button onclick="results()">Wild Card Search</button><div id="results"></div>'}
function results(){document.getElementById('results').innerHTML='<p>situs list count: 3</p><div class="k-grid" role="grid"><table><thead><tr>'+['No.','Frac','Dir','Street Name','Type','Dir','Unit-ID','Unit','Zip','Occupant'].map(x=>'<th>'+x+'</th>').join('')+'</tr></thead><tbody>'+[12105,12107,12109].map(n=>'<tr onclick="selected('+n+')">'+[n,'','','CRYSTALWOOD','DR','','','','77013','OWNER'].map(x=>'<td>'+x+'</td>').join('')+'</tr>').join('')+'</tbody></table></div>'}
function selected(n){window.selectedNumber=n;identity.innerHTML='<p>Situs: '+n+' CRYSTALWOOD DR 77013</p><p>Project: 123</p><button onclick="holds()">126 Situs Holds</button>';main.innerHTML='<h2>Project Master Maintenance</h2>'}
function grid(){return '<div class="k-grid" role="grid"><table><thead><tr>'+['Hold Date','Department','Hold Type','Action Status','Permit Type','Hold Reason','Hold By'].map(x=>'<th>'+x+'</th>').join('')+'</tr></thead><tbody>'+recordRows.map(row=>'<tr>'+row.map(x=>'<td>'+esc(x)+'</td>').join('')+'</tr>').join('')+'</tbody></table></div>'}
function holds(){main.innerHTML='<h2>(126) Situs Holds Associated</h2>'+grid()+'<button onclick="openHoldForm()">Add Additional Holds</button>'}
function openHoldForm(){main.innerHTML='<h2>(126) Situs Holds Associated</h2>'+grid()+'<h3>Enforce Additional Situs Holds</h3><div class="line"><label>Permit Type</label><input id="permit" onblur="permitDesc.value=this.value===\'GE\'?\'General Hold\':\'\'"><input class="wide" id="permitDesc" readonly></div><div class="line"><label>Department</label><input id="dept" onblur="deptDesc.value=this.value===\'140\'?\'PLANNING\':\'\'"><input class="wide" id="deptDesc" readonly></div><div class="line"><label>Hold Type</label><input id="holdCode" readonly><select onchange="holdCode.value=\'H\'"><option></option><option>HARD HOLD</option></select></div><div class="line"><label>Reason</label><input class="wide" id="reason" maxlength="80"></div><div class="line"><label>Authorized By</label><input readonly value="TEST"></div><button onclick="add()">Add</button>'}
function add(){main.insertAdjacentHTML('beforeend','<div role="dialog"><p>Process Add Request ?</p><button onclick="confirmedAdd()">Yes</button></div>')}
function confirmedAdd(){addClicks++;recordRows.push(['2026','PLANNING','HARD HOLD','','GE',reason.value,'TEST']);holds()}
search();</script></body></html>'''

PROJECT_EMPTY_FIXTURE=r''' 
window.recordRows=[];window.projectClicks=0;
function selected(n){window.selectedNumber=n;main.innerHTML='<table><thead><tr><th>Project</th><th>Group</th><th>Status</th></tr></thead><tbody><tr onclick="projectChosen()"><td>12053746</td><td>B</td><td>INACTIVE</td></tr></tbody></table>'}
function projectChosen(){projectClicks++;identity.innerHTML='<p>Situs: '+selectedNumber+' CRYSTALWOOD DR 77013</p><p>Project: 12053746</p><div role="treeitem" onclick="expandSitus()">Situs</div><div id="situsMenu"></div>';main.innerHTML='<h2>Project Master Maintenance</h2>'}
function expandSitus(){situsMenu.innerHTML='<div role="treeitem" onclick="holds()">126 Situs Holds</div>'}
function holds(){main.innerHTML='<h2>(126) Situs Holds Associated</h2>'+(recordRows.length?grid():'<p>NO Situs holds exist</p>')+'<button onclick="openHoldForm()">Add Additional Holds</button>'}
'''

MULTI_ADDRESS_FIXTURE=r''' 
window.sequence=[];window.byAddress={};
function holds(){main.innerHTML='<h2>(126) Situs Holds Associated</h2>'+(recordRows.length?grid():'<p>NO Situs holds exist</p>')+'<button onclick="openHoldForm()">Add Additional Holds</button>'}

const originalSearch=search,originalSelected=selected,originalConfirmedAdd=confirmedAdd;
search=function(){sequence.push('search');originalSearch()};
selected=function(n){recordRows=byAddress[n]||(byAddress[n]=[]);originalSelected(n)};
confirmedAdd=function(){sequence.push('add:'+selectedNumber);originalConfirmedAdd()};
'''


RETAINED_FIXTURE=r'''
window.projectClicks=0;window.savedAddresses=[];window.submittedZips=[];window.lastNumber=12105;
const doSearch=search,doResults=results,doSelected=selected,doConfirmedAdd=confirmedAdd;
search=function(){const header=identity.innerHTML;doSearch();identity.innerHTML=header;num.value=lastNumber;street.value='CRYSTALWOOD';zip.value='77013';zipPlus.value='0000';doResults();for(const tr of document.querySelectorAll('#results tbody tr'))if(Number(tr.cells[0].textContent)!==lastNumber)tr.remove()};
results=function(){submittedZips.push(zip.value);setTimeout(doResults,500)};
selected=function(n){window.pendingNumber=n;main.innerHTML='<h2>Find Situs</h2><table><thead><tr><th>Project</th><th>Group</th><th>Status</th></tr></thead><tbody><tr onclick="projectChosen()"><td>123</td><td>B</td><td>INACTIVE</td></tr></tbody></table>'};
projectChosen=function(){projectClicks++;doSelected(pendingNumber);lastNumber=pendingNumber};
confirmedAdd=function(){savedAddresses.push(selectedNumber);doConfirmedAdd();openHoldForm()};
'''

def main():return 0 if unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(ILMSTests)).wasSuccessful() else 1
if __name__=='__main__':raise SystemExit(main())
