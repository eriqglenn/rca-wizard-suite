"""Visible ILMS browser. Uses the same controls the planner uses; no private API."""
from __future__ import annotations
import copy
import os
import queue
import re
import threading
import time
from pathlib import Path
from urllib.parse import urlsplit

from .ilms_holds import Journal, address_parts, address_key, same_address, token, hold_reason, place_one

ILMS_URL='https://ilms.internal.example/Ilms_Client/login'
ILMS_HOSTS={'ilms.internal.example','ilms.internal.example'}

class ILMSError(ValueError):pass

def browser_error_message(error):
    message=str(error)
    if any(s in message.lower() for s in ('target page, context or browser has been closed','browser was closed','browser has been closed','connection closed')):
        return 'Chrome disconnected or was closed. Click Start / Resume to reopen ILMS and sign in. Confirmed holds are saved and will not be added again.'
    return message

class ILMSPage:
    def __init__(self,page,cancel=None):self.page=page;self.cancel=cancel;self.empty_holds=False

    def check(self):
        if self.cancel and self.cancel.is_set():raise InterruptedError('Stopped before the next hold.')
        if self.page.is_closed():raise ILMSError('The ILMS browser was closed. Open ILMS again.')
        if urlsplit(self.page.url).hostname not in ILMS_HOSTS:
            raise ILMSError('Sign into ILMS in the wizard browser, then try again.')

    def signed_in(self):
        # Do not read credentials. A visible authenticated navigation control
        # plus the absence of the login form is the only readiness signal.
        if urlsplit(self.page.url).hostname not in ILMS_HOSTS:return False
        if self.page.locator('input[type="password"]').filter(visible=True).count():return False
        return self.page.locator('nav').get_by_text('Address Search',exact=True).filter(visible=True).count()==1

    def wait_for_sign_in(self,seconds=900):
        self.wait(self.signed_in,'Sign-in was not completed. Click Start / Resume to try again.',seconds)

    def wait(self,predicate,description,seconds=45,cancellable=True):
        deadline=time.monotonic()+seconds;last=None
        while time.monotonic()<deadline:
            if cancellable:self.check()
            try:
                result=predicate()
                if result:return result
            except Exception as error:last=error
            self.page.wait_for_timeout(200)
        raise ILMSError(description+(f' ({last})' if last else ''))

    def exact(self,role,name):
        loc=self.page.get_by_role(role,name=re.compile(r'^\s*'+re.escape(name)+r'\s*$',re.I))
        return loc.filter(visible=True)

    def button(self,name):
        found=self.exact('button',name)
        if found.count()!=1:raise ILMSError(f'Could not identify one {name} button. Check the ILMS screen.')
        return found

    def click_text(self,name):
        if name=='Address Search':loc=self.page.locator('nav').get_by_text(name,exact=True).filter(visible=True)
        else:
            loc=self.exact('treeitem',name)
            if not loc.count():loc=self.page.get_by_text(name,exact=True).filter(visible=True)
        count=loc.count()
        # ILMS can expose the same navigation destination in both its common
        # actions and expanded Situs menu. This is not a property-row choice.
        if count>1 and name in ('126 Situs Holds','Situs','Address Search'):
            loc=loc.first
        elif count!=1:raise ILMSError(f'Could not identify {name}. Sign into ILMS and close any open dialog.')
        loc.click()

    def field(self,label,kind='input',occurrence=0,section=''):
        """Associate visible form labels with adjacent controls, not generated k-* IDs.

        ILMS has unassociated Kendo labels. Geometry is read from the live DOM,
        constrained to the named form, and ambiguity fails before a submission.
        """
        selector='[role="combobox"],select' if kind=='combo' else 'input:not([type="hidden"]),textarea'
        result=self.page.evaluate('''({label,selector,occurrence,section})=>{
          const visible=e=>{const r=e.getBoundingClientRect();const s=getComputedStyle(e);return r.width>0&&r.height>0&&s.visibility!=='hidden'&&s.display!=='none'};
          const norm=s=>(s||'').replace(/[\\u200b-\\u200d\\ufeff]/g,'').replace(/\\s+/g,' ').trim().replace(/:$/,'').toLowerCase();
          let floor=-Infinity;
          if(section){const h=[...document.querySelectorAll('h1,h2,h3,h4,h5,h6,legend')].filter(e=>visible(e)&&norm(e.innerText)===norm(section));
            if(h.length!==1)return {error:'Expected one form heading: '+section};floor=h[0].getBoundingClientRect().bottom;}
          // Form captions can be bare table-cell text or Kendo elements.
          // Exclude data grids, not layout tables, and measure the caption text
          // itself when its element also contains a control.
          const labels=[...document.querySelectorAll('body *')].filter(e=>visible(e)&&!e.closest('.k-grid,[role="grid"],input,textarea,select,option,button,[role="combobox"]')&&norm(e.innerText)===norm(label)&&e.getBoundingClientRect().top>=floor-2
            &&![...e.children].some(c=>visible(c)&&norm(c.innerText)===norm(label))).sort((a,b)=>a.getBoundingClientRect().top-b.getBoundingClientRect().top||a.getBoundingClientRect().left-b.getBoundingClientRect().left);
          if(!labels[occurrence])return {error:'Missing label: '+label};
          const caption=labels[occurrence];let l=caption.getBoundingClientRect();
          const textNodes=[...caption.childNodes].filter(n=>n.nodeType===Node.TEXT_NODE&&norm(n.textContent)===norm(label));
          if(textNodes.length===1){const range=document.createRange();range.selectNodeContents(textNodes[0]);const r=range.getBoundingClientRect();if(r.width&&r.height)l=r;}
          const fields=[...document.querySelectorAll(selector)];
          const ranked=fields.map((e,index)=>({e,index,r:e.getBoundingClientRect()})).filter(o=>visible(o.e)&&o.r.top>=floor-2&&!o.e.closest('.k-grid,[role="grid"]')).map(o=>{
            const same=Math.abs((o.r.top+o.r.height/2)-(l.top+l.height/2))<18&&o.r.left>=l.right-5&&o.r.left-l.right<250;
            const below=o.r.top>=l.bottom-3&&o.r.top-l.bottom<48&&Math.abs(o.r.left-l.left)<60;
            return {...o,score:same?Math.max(0,o.r.left-l.right)+Math.abs(o.r.top-l.top)*3:below?Math.abs(o.r.left-l.left)+Math.max(0,o.r.top-l.bottom)*2:Infinity};
          }).filter(o=>Number.isFinite(o.score)).sort((a,b)=>a.score-b.score);
          if(!ranked.length)return {error:'No control beside '+label};
          if(ranked[1]&&Math.abs(ranked[1].score-ranked[0].score)<1)return {error:'Ambiguous control beside '+label};
          return {index:ranked[0].index};
        }''',{'label':label,'selector':selector,'occurrence':occurrence,'section':section})
        if 'error' in result:raise ILMSError(result['error'])
        return self.page.locator(selector).nth(result['index'])

    def fill(self,label,value,section=''):
        control=self.field(label,section=section)
        if not control.is_editable():raise ILMSError(label+' is not editable on this screen.')
        limit=control.get_attribute('maxlength')
        if limit and int(limit)>0 and len(value)>int(limit):raise ILMSError(label+' is too long for ILMS; it was not truncated.')
        control.fill(value);control.press('Tab')
        if control.input_value().strip().upper()!=value.strip().upper():raise ILMSError(label+' did not retain the entered value.')

    def choose(self,label,choices,occurrence=0,section=''):
        combo=self.field(label,kind='combo',occurrence=occurrence,section=section)
        if combo.evaluate('(e)=>e.tagName')=='SELECT':
            options=combo.locator('option');matches=[i for i in range(options.count()) if token(options.nth(i).inner_text()) in choices]
            if len(matches)!=1:raise ILMSError('Could not identify one '+label+' choice.')
            combo.select_option(index=matches[0]);return
        combo.click()
        options=self.page.get_by_role('option').filter(visible=True)
        self.wait(lambda:options.count()>0,'No '+label+' options appeared.',15)
        matches=[i for i in range(options.count()) if token(options.nth(i).inner_text()) in choices]
        if len(matches)!=1:
            self.page.keyboard.press('Escape');raise ILMSError('Could not identify one '+label+' choice.')
        options.nth(matches[0]).scroll_into_view_if_needed();options.nth(matches[0]).click()

    def situs(self):
        text=self.page.locator('body').inner_text()
        m=re.search(r'\bSitus:\s*(.*?)\s*(?:\bProject:|\bApply Permit Use:)',text,re.S|re.I)
        if not m:raise ILMSError('ILMS is not showing a selected situs address.')
        return re.sub(r'\s+',' ',m[1]).strip()

    def assert_address(self,address):
        self.check();current=self.situs()
        if not same_address(current,address):
            raise ILMSError(f'ILMS shows {current}; expected {address}. No hold will be placed on a different address.')

    def grids(self):return self.page.locator('[role="grid"],.k-grid,table').filter(visible=True)

    def grid(self,required):
        grids=self.grids();matches=[]
        for i in range(grids.count()):
            g=grids.nth(i);heads=g.locator('th,[role="columnheader"]').all_inner_texts()
            if not heads:heads=g.get_by_role('row').first.locator('td,[role="gridcell"]').all_inner_texts()
            if all(any(token(h)==token(w) for h in heads) for w in required):
                if not any(parent.evaluate('(e,child)=>e.contains(child)',g.element_handle()) for parent in matches):matches.append(g)
        # Some grids have a separate header and body table; use the Kendo wrapper.
        if len(matches)==1:
            g=matches[0]
            if not g.locator('tbody tr,[role="row"]').count():raise ILMSError('ILMS grid has no rows.')
            return g
        raise ILMSError('Could not identify the ILMS '+', '.join(required)+' table.')

    def rows(self,grid):
        rows=grid.locator('tr,[role="row"]');out=[]
        for i in range(rows.count()):
            row=rows.nth(i);cells=row.locator('td,[role="gridcell"]')
            if cells.count():out.append((row,[re.sub(r'\s+',' ',v).strip() for v in cells.all_inner_texts()]))
        return out

    def find(self,address):
        self.check();parts=address_parts(address)
        if not address_key(address):raise ILMSError('A complete numbered property address is required.')
        self.click_text('Address Search')
        self.wait(lambda:self.page.get_by_role('heading',name='Find Situs',exact=True).is_visible(),'Address Search did not open.')
        for label in ('Number','Street Name','Unit Type','Unit No.','Zip Code','Zip Plus'):
            self.fill(label,{'Number':parts['number'],'Street Name':parts['street']}.get(label,''),section='Address Search')
        # Starting each search from a fresh form prevents prior directions/units leaking.
        from .ilms_holds import STREET_TYPES, DIRECTIONS
        self.choose('Fraction',{token(parts['fraction'])},section='Address Search')
        self.choose('Direction',{parts['prefix']}|{token(k) for k,v in DIRECTIONS.items() if v==parts['prefix']},section='Address Search')
        types={token(k) for k,v in STREET_TYPES.items() if v==parts['type']}
        # Prefer full display wording (DRIVE) over obsolete duplicate codes (DR).
        self.choose('Street Type',types or {parts['type']},section='Address Search')
        self.choose('Direction',{parts['suffix']}|{token(k) for k,v in DIRECTIONS.items() if v==parts['suffix']},1,section='Address Search')
        # ILMS keeps the last result table when Address Search reopens. Arm an
        # observer before clicking so the old count/rows cannot satisfy readiness.
        self.page.evaluate('''()=>{
          if(window.__rcaSearchObserver)window.__rcaSearchObserver.disconnect();
          window.__rcaSearchChanged=false;
          const isResult=n=>n.nodeType===1&&(n.matches('.k-grid,[role="grid"]')||n.querySelector('.k-grid,[role="grid"]')||/situs list count:/i.test(n.textContent||''));
          window.__rcaSearchObserver=new MutationObserver(records=>{
            if(records.some(r=>(r.target.nodeType===1?r.target:r.target.parentElement)?.closest('.k-grid,[role="grid"]')||[...r.addedNodes,...r.removedNodes].some(isResult))){window.__rcaSearchChanged=true;}
          });
          window.__rcaSearchObserver.observe(document.body,{childList:true,subtree:true,characterData:true});
        }''')
        try:
            self.button('Wild Card Search').click()
            def fresh_results():
                if not self.page.evaluate('window.__rcaSearchChanged===true'):return False
                if self.page.locator('.k-loading-mask,[aria-busy="true"]').filter(visible=True).count():return False
                count=self.page.get_by_text(re.compile(r'situs list count:\s*\d+',re.I))
                if not count.count():return False
                self.grid(['Street Name','Type','Zip'])
                return True
            self.wait(fresh_results,'No fresh address list returned. Check the address and ILMS connection.')
        finally:self.page.evaluate('()=>{if(window.__rcaSearchObserver)window.__rcaSearchObserver.disconnect();}')
        grid=self.grid(['Street Name','Type','Zip']);headers=[token(v) for v in grid.locator('th,[role="columnheader"]').all_inner_texts()]
        # Headers have two Dir columns and sometimes a leading selection column.
        if 'NO' not in headers or 'STREETNAME' not in headers:raise ILMSError('Unrecognized address table headers.')
        def cell(values,name,occurrence=0):
            indexes=[i for i,h in enumerate(headers) if h==token(name)]
            return values[indexes[occurrence]] if len(indexes)>occurrence and indexes[occurrence]<len(values) else ''
        matches={};last_scroll=None
        for _ in range(200):
            self.check()
            for row,values in self.rows(grid):
                candidate=' '.join(v for v in (cell(values,'No.'),cell(values,'Frac'),cell(values,'Dir'),cell(values,'Street Name'),cell(values,'Type'),cell(values,'Dir',1)) if v)
                unit=cell(values,'Unit');unit_id=cell(values,'Unit-ID')
                if unit or unit_id:
                    marker=unit_id if unit_id.upper() in ('APT','UNIT','STE','BLDG','#') else 'UNIT'
                    candidate+=' '+marker+' '+(' '.join(v for v in ((unit_id if marker=='UNIT' and unit_id!='UNIT' else ''),unit) if v))
                if same_address(candidate,address):matches[tuple(values)]=candidate
            pos=grid.evaluate('''e=>{const all=[e,...e.querySelectorAll('*')];const s=all.find(x=>x.scrollHeight>x.clientHeight+5&&x.clientHeight>50&&/(auto|scroll)/.test(getComputedStyle(x).overflowY));if(!s)return null;const before=s.scrollTop;s.scrollTop=Math.min(s.scrollTop+Math.max(80,s.clientHeight-40),s.scrollHeight);return [s.scrollTop,s.scrollHeight,s.clientHeight,before]}''')
            if not pos or pos[0]==pos[3] or pos==last_scroll:break
            last_scroll=pos;self.page.wait_for_timeout(150)
        if len(matches)!=1:raise ILMSError(f'{len(matches)} exact situs matches for {address}. Check its street direction, fraction and unit; no hold was placed.')
        wanted=next(iter(matches));candidate=matches[wanted]
        # Re-find the chosen row after virtual scrolling, never reuse a stale row handle.
        grid.evaluate('''e=>{for(const x of [e,...e.querySelectorAll('*')])if(x.scrollHeight>x.clientHeight+5&&x.clientHeight>50&&/(auto|scroll)/.test(getComputedStyle(x).overflowY))x.scrollTop=0}''')
        clicked=False
        for _ in range(200):
            for row,values in self.rows(grid):
                if tuple(values)==wanted:row.scroll_into_view_if_needed();row.click();clicked=True;break
            if clicked:break
            moved=grid.evaluate('''e=>{const s=[e,...e.querySelectorAll('*')].find(x=>x.scrollHeight>x.clientHeight+5&&x.clientHeight>50&&/(auto|scroll)/.test(getComputedStyle(x).overflowY));if(!s)return false;const p=s.scrollTop;s.scrollTop+=Math.max(80,s.clientHeight-40);return s.scrollTop>p}''')
            if not moved:break
            self.page.wait_for_timeout(150)
        if not clicked:raise ILMSError('The matched ILMS row is no longer available. Search again.')
        def opened_or_projects():
            try:
                g=self.grid(['Project','Group','Status'])
                if self.rows(g):return 'projects'
            except ILMSError:pass
            if not self.page.get_by_role('heading',name='Find Situs',exact=True).is_visible() and self._address_visible(address):return 'situs'
            return None
        destination=self.wait(opened_or_projects,'The selected address did not open a situs or project list.')
        if destination=='projects':
            # The list was opened from the one exact situs match. A project
            # is only the navigation route to its situs-level holds.
            rows=self.rows(self.grid(['Project','Group','Status']))
            if not rows:raise ILMSError('No project is available for this situs address.')
            rows[0][0].scroll_into_view_if_needed();rows[0][0].click()
        self.wait(lambda:not self.page.get_by_role('heading',name='Find Situs',exact=True).is_visible() and self._address_visible(address),'The selected situs did not match the requested address.')
        self.open_holds();return {'address':self.situs(),'search_row':list(wanted),'url':self.page.url}

    def _address_visible(self,address):
        try:return same_address(self.situs(),address)
        except ILMSError:return False

    def open_holds(self):
        self.empty_holds=False
        def holds_link():
            return self.exact('treeitem','126 Situs Holds').count() or self.page.get_by_text('126 Situs Holds',exact=True).filter(visible=True).count()
        # The identity header can arrive before the sidebar during navigation.
        self.wait(lambda:holds_link() or self.exact('treeitem','Situs').count() or self.page.get_by_text('Situs',exact=True).filter(visible=True).count(),
                  'The ILMS navigation menu did not load. Resume after checking the project screen.',15)
        if not holds_link():
            self.click_text('Situs')
            self.wait(holds_link,'Expand the Situs menu to locate 126 Situs Holds.')
        self.click_text('126 Situs Holds')
        self.wait(lambda:self.page.get_by_role('heading',name=re.compile(r'\(126\) Situs Holds Associated')).is_visible(),'126 Situs Holds did not open.')
        self.wait(lambda:self.page.get_by_role('button',name='Add Additional Holds',exact=True).is_visible(),'The hold list is not ready.')
        def loaded():
            try:self.grid(['Hold Date','Department','Hold Type','Hold Reason']);return True
            except ILMSError:
                if self.page.get_by_text('NO Situs holds exist',exact=True).filter(visible=True).count():
                    self.empty_holds=True;return True
                return False
        self.wait(loaded,'ILMS did not return a hold list or confirm that no holds exist.')

    def holds(self):
        try:grid=self.grid(['Hold Date','Department','Hold Type','Hold Reason'])
        except ILMSError:
            if self.empty_holds:return []
            raise
        self.empty_holds=False
        return [values for _,values in self.rows(grid) if any(values) and 'No records available' not in ' '.join(values)]

    @staticmethod
    def matching_count(rows,reason):
        return sum(1 for values in rows if any(v.strip().casefold()==reason.casefold() for v in values)
                   and 'GE' in values and any(token(v)=='PLANNING' for v in values) and any(token(v)=='HARDHOLD' for v in values))

    def prepare(self,reason):
        self.check();self.button('Add Additional Holds').click()
        self.wait(lambda:self.page.get_by_role('heading',name='Enforce Additional Situs Holds',exact=True).is_visible(),'The additional-hold form did not open.')
        self.fill('Permit Type','GE',section='Enforce Additional Situs Holds')
        self.fill('Department','140',section='Enforce Additional Situs Holds')
        self.choose('Hold Type',{'HARDHOLD','HHARDHOLD'},section='Enforce Additional Situs Holds')
        self.fill('Reason',reason,section='Enforce Additional Situs Holds')
        for label,expected in (('Permit Type','GE'),('Department','140'),('Hold Type','H'),('Reason',reason)):
            actual=self.field(label,section='Enforce Additional Situs Holds').input_value().strip()
            if actual.casefold()!=expected.casefold():raise ILMSError(label+' does not match the intended hold. Nothing submitted.')
        # Date and read-only Authorized By are supplied by ILMS for the signed-in planner.

    def submit(self):
        self.button('Add').click(timeout=15000)
        dialog=self.page.get_by_role('dialog').filter(has_text=re.compile(r'Process Add Request\s*\?',re.I)).filter(visible=True)
        self.wait(lambda:dialog.count()==1,'ILMS did not show the expected Add confirmation.',20,cancellable=False)
        yes=dialog.get_by_role('button',name='Yes',exact=True)
        if yes.count()!=1:raise ILMSError('Could not identify the Add confirmation button. Check ILMS before retrying.')
        yes.click(timeout=15000)

    def verify_added(self,reason,before_count):
        def verified():
            rows=self.holds()
            return rows if self.matching_count(rows,reason)==before_count+1 else None
        # Once Add was clicked, allow verification to finish even if Stop is pressed.
        return self.wait(verified,'The additional hold could not be verified in the ILMS list.',60,cancellable=False)

class ILMSBrowser:
    def __init__(self,events,profile=None):
        self.events=events;self.commands=queue.Queue();self.stop=threading.Event();self.closed=threading.Event();self.thread=None
        self.profile=Path(profile) if profile else Path(os.environ.get('LOCALAPPDATA',Path.home()))/'RCAWizardSuite'/'ILMSBrowser'

    def start(self):
        if self.thread and self.thread.is_alive():return
        # A failed launch may leave an unconsumed command from the prior run.
        while True:
            try:self.commands.get_nowait()
            except queue.Empty:break
        self.closed.clear();self.stop.clear();self.thread=threading.Thread(target=self._run,daemon=True);self.thread.start()

    def send(self,kind,data):
        if not self.thread or not self.thread.is_alive():raise ILMSError('Open ILMS and sign in first.')
        self.stop.clear();self.commands.put((kind,copy.deepcopy(data)))

    def close(self):self.stop.set();self.closed.set()

    def process_batch(self,adapter,kind,data,journal):
        """Stop on the first failed operation; later addresses were not checked.

        In particular, never turn a lost session or a missing navigation menu
        into a series of property failures. Journal reconciliation still owns
        the decision to retry any Add that may already have reached ILMS.
        """
        paused=''
        for index,p in enumerate(data['properties']):
            if not p.get('selected'):continue
            try:
                if self.stop.is_set():raise InterruptedError('Stopped before the next hold.')
                self.events.put(('progress','Processing '+p['address']+'…'))
                if kind=='find':
                    match=adapter.find(p['address']);holds=adapter.holds()
                    result={'status':'Address matched','detail':f'{len(holds)} existing hold(s); a new application hold will still be added.','match':match}
                else:result=place_one(adapter,p,hold_reason(data['application']),journal,self.stop)
                p.update(result);self.events.put(('property',copy.deepcopy(p)))
            except InterruptedError:
                paused='Stopped before the next hold.'
                p.update(status='Paused',detail=paused);self.events.put(('property',copy.deepcopy(p)));break
            except Exception as error:
                paused=browser_error_message(error)
                state=(journal.entry(p.get('key','')) or {}).get('state')
                p.update(status='Needs review' if state in ('submitting','uncertain') else 'Paused',detail=paused)
                self.events.put(('property',copy.deepcopy(p)));break
        if paused:
            for remaining in data['properties'][index+1:]:
                if not remaining.get('selected'):continue
                state=(journal.entry(remaining.get('key','')) or {}).get('state')
                if state=='added':
                    remaining.update(status='Completed in this batch',detail='Previously verified hold retained.')
                elif state in ('submitting','uncertain'):
                    remaining.update(status='Needs review',detail='Previous Add requires reconciliation before retrying.')
                else:
                    remaining.update(status='Not attempted',detail='Batch paused before this address. Click Start / Resume to continue.')
                self.events.put(('property',copy.deepcopy(remaining)))
        report=journal.report(data['properties'])
        completed=sum(p.get('status') in ('Hold added','Completed in this batch','Confirmed after interruption') for p in data['properties'])
        review=sum(p.get('status') in ('Needs review','Needs address') for p in data['properties'])
        pending=sum(p.get('selected') and p.get('status') not in ('Hold added','Completed in this batch','Confirmed after interruption','Address matched') for p in data['properties'])
        message=('Paused. ' if paused else 'Finished. ')+f'{completed} completed; {pending} remaining; {review} need review. '
        if paused:message+=paused+' Resume keeps confirmed holds and checks uncertain submissions before adding. '
        self.events.put(('done',{'message':message+'Report: '+str(report),'report':str(report)}))

    def _run(self):
        try:
            from playwright.sync_api import sync_playwright
            with sync_playwright() as runtime:
                with runtime.chromium.launch_persistent_context(str(self.profile),channel='chrome',headless=False,
                        viewport={'width':1500,'height':950},accept_downloads=False) as context:
                    page=context.pages[0] if context.pages else context.new_page();page.set_default_timeout(15000)
                    page.goto(ILMS_URL,wait_until='domcontentloaded',timeout=60000)
                    self.events.put(('browser_ready','Chrome is open. Sign in to ILMS; the wizard will continue automatically.'))
                    adapter=ILMSPage(page,self.stop)
                    while not self.closed.is_set() and not page.is_closed():
                        try:kind,data=self.commands.get_nowait()
                        except queue.Empty:page.wait_for_timeout(150);continue
                        journal=None
                        try:
                            if kind=='automatic':
                                self.events.put(('progress','Waiting for you to sign in to ILMS in Chrome…'))
                                adapter.wait_for_sign_in()
                                from .workbooks import read_shared_bytes
                                import hashlib
                                if hashlib.sha256(read_shared_bytes(data['workbook'])).hexdigest()!=data['workbook_sha256']:
                                    raise ILMSError('The calculation sheet changed during sign-in. Start again to reload the properties.')
                                self.events.put(('progress','Signed in. Finding properties and placing application holds…'))
                            journal=Journal(data['output'],data['wizard'],data['application']);journal.acquire()
                            self.process_batch(adapter,kind,data,journal)
                        except Exception as error:self.events.put(('error',browser_error_message(error)))
                        finally:
                            if journal:journal.close()
        except Exception as error:self.events.put(('error',browser_error_message(error)))
        finally:self.events.put(('browser_closed',None))
