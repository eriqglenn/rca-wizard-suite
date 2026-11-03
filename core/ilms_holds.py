"""126 situs hold plans and durable, case-specific submission history."""
from __future__ import annotations
import csv
import hashlib
import io
import json
import os
import re
from datetime import datetime, timezone
from pathlib import Path

from .workbooks import read_shared_bytes

WIZARDS = ('mls', 'mbl', 'mls_renewal', 'mbl_renewal')
STREET_TYPES = {'STREET':'ST','DRIVE':'DR','AVENUE':'AVE','ROAD':'RD','BOULEVARD':'BLVD',
                'LANE':'LN','COURT':'CT','CIRCLE':'CIR','PLACE':'PL','PARKWAY':'PKWY',
                'FREEWAY':'FWY','HIGHWAY':'HWY','TERRACE':'TER','TRAIL':'TRL','WAY':'WAY',
                'BEND':'BND','LOOP':'LOOP','SQUARE':'SQ','ALLEY':'ALY'}
DIRECTIONS = {'NORTH':'N','SOUTH':'S','EAST':'E','WEST':'W','NORTHEAST':'NE',
              'NORTHWEST':'NW','SOUTHEAST':'SE','SOUTHWEST':'SW'}

def clean(value):
    if value is None:return ''
    if isinstance(value,float) and value.is_integer():return str(int(value))
    return re.sub(r'\s+',' ',str(value)).strip()

def token(value):
    return re.sub(r'[^A-Z0-9]','',clean(value).upper())

def application_number(value):
    value=clean(value).upper()
    value=re.sub(r'^(?:MLS|MBL)\s*#?\s*','',value)
    value=re.sub(r'\s+','',value)
    if not re.fullmatch(r'\d+[A-Z0-9-]*',value):
        raise ValueError('Enter the application number, including REN for a renewal (for example 187REN).')
    return value

def hold_reason(application):
    return f'MLS/MBL {application_number(application)} call planning 6637'

def address_parts(address):
    text=clean(address).upper().replace('.','').replace(',',' ')
    text=re.sub(r'\s+(?:HOUSTON\s+)?(?:TX\s+)?\d{5}(?:-\d{4})?$','',text).strip()
    words=text.split();result={'number':'','fraction':'','prefix':'','street':'','type':'','suffix':'','unit_type':'','unit':''}
    if not words or not re.fullmatch(r'[1-9]\d*[A-Z]?',words[0]):return result
    result['number']=words.pop(0)
    if words and re.fullmatch(r'\d+/\d+',words[0]):result['fraction']=words.pop(0)
    directions=set(DIRECTIONS.values())
    if words and DIRECTIONS.get(words[0],words[0]) in directions:result['prefix']=DIRECTIONS.get(words[0],words[0]);words.pop(0)
    units={'APT','APARTMENT','UNIT','STE','SUITE','BLDG','BUILDING','#'}
    unit_at=next((i for i,w in enumerate(words) if w in units or w.startswith('#')),None)
    if unit_at is not None:
        marker=words[unit_at];result['unit_type']={'APARTMENT':'APT','SUITE':'STE','BUILDING':'BLDG'}.get(marker,marker)
        result['unit']=' '.join(words[unit_at+1:]) if not marker.startswith('#') or marker=='#' else marker[1:]+' '+' '.join(words[unit_at+1:])
        result['unit']=result['unit'].strip();words=words[:unit_at]
    if words and DIRECTIONS.get(words[-1],words[-1]) in directions:
        last=words.pop();result['suffix']=DIRECTIONS.get(last,last)
    if words:
        last=STREET_TYPES.get(words[-1],words[-1])
        if last in STREET_TYPES.values():result['type']=last;words.pop()
    result['street']=' '.join(words)
    return result

def address_key(address):
    p=address_parts(address)
    if not p['number'] or not p['street'] or not p['type']:return ''
    # A unit identifier remains significant; common UNIT/APT/# notation is equivalent.
    return '|'.join(token(p[k]) for k in ('number','fraction','prefix','street','type','suffix','unit'))

def same_address(left,right):
    a,b=address_key(left),address_key(right)
    return bool(a and a==b)

def load_plan(path,wizard):
    """Read property addresses only; EXC and repeated legal-lot rows are retained/grouped."""
    from openpyxl import load_workbook
    if wizard not in WIZARDS:raise ValueError('Unknown wizard type.')
    raw=read_shared_bytes(path);book=load_workbook(io.BytesIO(raw),read_only=True,data_only=True)
    try:
        s1=next((s for s in book if s.title.lower().startswith('step 1')),None)
        s2=next((s for s in book if s.title.lower().startswith('step 2')),None)
        if s1 is None:raise ValueError('The calculation sheet needs a Step 1 property list.')
        headings={token(v):i for i,v in enumerate(next(s1.iter_rows(max_row=1,values_only=True))) if v}
        if not any(k in headings for k in ('SITEADDR1','SITEADDRESS','ADDRESS','STRNAME')):
            raise ValueError('Step 1 does not contain property-address columns.')
        ci={token(r[0]):r[1] for r in book['Case Info'].iter_rows(max_col=2,values_only=True) if r[0]} if 'Case Info' in book else {}
        app=clean(s2['B2'].value) if s2 else ''
        app=app or next((clean(ci[k]) for k in ('APPLICATIONNO','APPLICATIONNUMBER','FILENUMBER','CASENUMBER') if ci.get(k)),'')
        grouped={};entries=[]
        for rn,row in enumerate(s1.iter_rows(min_row=2,max_row=min(s1.max_row,10000),values_only=True),2):
            def get(*names):
                return next((clean(row[headings[token(n)]]) for n in names if token(n) in headings and headings[token(n)]<len(row) and clean(row[headings[token(n)]])), '')
            account=get('TAX_ID','ALT_TAX_ID','HCAD_NUM','HCAD_NUM_1')
            address=get('Site_addr_1','Site Address','Address')
            legal=get('Legal_Dscr_1','Legal Description','Property Description')
            if not address:address=' '.join(get(k) for k in ('Str_num','Str_pfx','Str_name','Str_sfx') if get(k))
            if not any((account,address,legal)):continue
            unit=get('str_unit')
            if unit and unit not in ('0','0.0') and not address_parts(address)['unit']:address+=' UNIT '+unit
            key=address_key(address)
            merge_key=key or 'unmatched:'+str(rn)
            if merge_key in grouped:
                entry=grouped[merge_key];entry['rows'].append(rn)
                if account and account not in entry['accounts']:entry['accounts'].append(account)
                continue
            entry={'id':hashlib.sha256(merge_key.encode()).hexdigest()[:20], 'address':address,
                   'key':key,'rows':[rn],'accounts':[account] if account else [],'land_use':get('Land Use'),
                   'status':'Ready to find' if key else 'Needs address','selected':bool(key),
                   'detail':'' if key else 'No complete property address. Enter the actual ILMS address; 0 is not a situs match.'}
            entries.append(entry);grouped[merge_key]=entry
        if not entries:raise ValueError('Step 1 has no properties to process.')
        return {'wizard':wizard,'application':app,'workbook':str(Path(path).resolve()),
                'workbook_sha256':hashlib.sha256(raw).hexdigest(),'properties':entries}
    finally:book.close()

def now():return datetime.now(timezone.utc).isoformat(timespec='seconds')

class Journal:
    """Persist intent before Add. An interrupted attempt is never retried blindly."""
    def __init__(self,folder,wizard,application):
        app=application_number(application)
        if wizard not in WIZARDS:raise ValueError('Unknown wizard type.')
        self.path=Path(folder)/f'ILMS_Holds_{wizard}_{app}.json'
        self.path.parent.mkdir(parents=True,exist_ok=True)
        self.data={'schema':1,'wizard':wizard,'application':app,'reason':hold_reason(app),'entries':{},'created':now()}
        if self.path.exists():
            data=json.loads(self.path.read_text(encoding='utf8'))
            if data.get('schema')!=1 or data.get('wizard')!=wizard or data.get('application')!=app:
                raise ValueError('The existing ILMS log has a different case. Choose the correct case folder.')
            self.data=data
        self.lock_path=self.path.with_suffix('.lock');self.lock_handle=None

    def acquire(self):
        # OS file locking releases on a crash; a stale lock file never blocks a retry.
        import msvcrt
        self.lock_handle=self.lock_path.open('a+b');self.lock_handle.seek(0)
        if self.lock_path.stat().st_size==0:self.lock_handle.write(b'0');self.lock_handle.flush()
        self.lock_handle.seek(0)
        try:msvcrt.locking(self.lock_handle.fileno(),msvcrt.LK_NBLCK,1)
        except OSError:
            self.lock_handle.close();self.lock_handle=None
            raise ValueError('Another wizard is placing holds for this case. Wait for it to finish.')
        if self.path.exists():self.data=json.loads(self.path.read_text(encoding='utf8'))

    def close(self):
        if self.lock_handle:
            import msvcrt
            self.lock_handle.seek(0);msvcrt.locking(self.lock_handle.fileno(),msvcrt.LK_UNLCK,1)
            self.lock_handle.close();self.lock_handle=None

    def save(self):
        self.data['updated']=now();temporary=self.path.with_suffix('.tmp')
        with temporary.open('w',encoding='utf8') as stream:
            json.dump(self.data,stream,indent=2,ensure_ascii=False);stream.flush();os.fsync(stream.fileno())
        os.replace(temporary,self.path)

    def entry(self,key):return self.data['entries'].get(key)

    def record(self,key,**values):
        self.data['entries'][key]={**self.data['entries'].get(key,{}),**values,'updated':now()};self.save()

    def report(self,properties):
        path=self.path.with_suffix('.csv');tmp=path.with_suffix('.csv.tmp')
        with tmp.open('w',newline='',encoding='utf-8-sig') as stream:
            writer=csv.writer(stream);writer.writerow(['Application','Property address','Step 1 rows','Result','Details','Reason'])
            for p in properties:
                row=[self.data['application'],p['address'],', '.join(map(str,p['rows'])),p['status'],p.get('detail',''),self.data['reason']]
                writer.writerow(["'"+v if isinstance(v,str) and v.startswith(('=','+','-','@')) else v for v in row])
        os.replace(tmp,path);return path

def place_one(adapter,property,reason,journal,cancel=None):
    """Only a caller's explicitly selected property may be submitted."""
    if not property.get('selected'):return {'status':'Not selected','detail':''}
    key=property.get('key')
    if not key:raise ValueError('Select a complete property address before placing a hold.')
    previous=journal.entry(key)
    if previous and previous.get('state')=='added':
        return {'status':'Completed in this batch','detail':'This case log already records a verified addition. Older ILMS holds do not otherwise block a new hold.'}
    if cancel and cancel.is_set():raise InterruptedError('Stopped before the next hold.')
    match=adapter.find(property['address'])
    if not same_address(match['address'],property['address']):raise ValueError('ILMS address changed; nothing submitted.')
    before=adapter.holds()
    count=adapter.matching_count(before,reason)
    if previous and previous.get('state') in ('submitting','uncertain'):
        if count==previous['before_count']+1:
            journal.record(key,state='added',confirmation='Matching additional hold found after interrupted attempt',after=before)
            return {'status':'Confirmed after interruption','detail':'The additional matching hold is present. No second hold was submitted.'}
        raise ValueError('Previous submission is uncertain. Inspect the ILMS hold list and case log before retrying; the wizard will not submit it again automatically.')
    adapter.prepare(reason)
    adapter.assert_address(property['address'])
    if cancel and cancel.is_set():raise InterruptedError('Stopped before Add.')
    journal.record(key,state='submitting',address=property['address'],reason=reason,before_count=count,before=before,match=match)
    try:
        # Never repeat this click after a timeout. The journal makes a restart safe.
        adapter.submit()
        after=adapter.verify_added(reason,count)
        adapter.assert_address(property['address'])
        journal.record(key,state='added',after=after,confirmation='Additional hold verified in ILMS')
    except Exception as error:
        journal.record(key,state='uncertain',error=str(error))
        raise ValueError('Add was attempted, but confirmation is incomplete. Inspect ILMS before retrying. '+str(error)) from error
    return {'status':'Hold added','detail':reason}
