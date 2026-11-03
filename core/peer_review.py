"""Evidence-based peer review of the actual saved case files. Advisory, read-only."""
from __future__ import annotations
import hashlib
import html
import json
import re
import tempfile
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path

from .peer_readers import clean, norm, number, read_calculation, read_document, hcad_record, check_cancel
from .workbooks import read_shared_bytes

VERSION = 1
ELIGIBLE = {'SFR','MF','COM','VAC'}


def show(value):
    return f'{value:,.2f}'.rstrip('0').rstrip('.') if isinstance(value, float) else str(value)


def address_key(value):
    text = clean(value).upper()
    for long, short in [('STREET','ST'),('AVENUE','AVE'),('ROAD','RD'),('DRIVE','DR'),('BOULEVARD','BLVD'),('LANE','LN'),('NORTH','N'),('SOUTH','S'),('EAST','E'),('WEST','W')]:
        text = re.sub(r'\b'+long+r'\b', short, text)
    return norm(text)


def words_number(text):
    names = 'zero one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen sixteen seventeen eighteen nineteen'.split()
    mapping = {w:i for i,w in enumerate(names)}
    mapping.update(dict(zip('twenty thirty forty fifty sixty seventy eighty ninety'.split(), range(20,100,10))))
    total = current = 0
    tokens = text.lower().replace('-', ' ').split()
    if not tokens:
        return None
    for word in tokens:
        if word == 'and': continue
        if word in mapping: current += mapping[word]
        elif word == 'hundred': current = max(current,1)*100
        elif word in ('thousand','million'):
            total += max(current,1)*(1000 if word=='thousand' else 1000000); current=0
        else: return None
    return total+current


WORD = r'(?:zero|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety|hundred|thousand|million|and)'
WORD_NUMBER = rf'\b({WORD}(?:[ -]+{WORD})*)\s*\(([\d,]+)\)'


class Review:
    def __init__(self, wizard):
        self.result = {'schema':VERSION, 'created_utc':datetime.now(timezone.utc).isoformat(),
                       'wizard':wizard, 'files':[], 'findings':[], 'facts':{}, 'status':'complete'}

    def add(self, severity, title, details, locations=(), source=None, expected=None, actual=None):
        if source and source.get('method') in ('OCR','partial','unreadable'):
            if severity in ('Error','Passed'): severity='Needs review'
            details += ' Confirm the text against the page image; this comparison uses recognized or partial text.'
        self.result['findings'].append({'severity':severity, 'title':title, 'details':details,
                                      'locations':list(locations), 'expected':expected, 'actual':actual})

    def compare(self, title, expected, actual, locations, source=None, tolerance=0):
        same = abs(expected-actual)<=tolerance if isinstance(expected,(float,int)) and isinstance(actual,(float,int)) else expected==actual
        self.add('Passed' if same else 'Error', title,
                 f'Expected {show(expected)}; found {show(actual)}.', locations, source, expected, actual)


def calculation_checks(review, facts):
    s2 = facts['step2']; table=facts['table']; rows=facts['rows']; minimum=facts['minimum']
    for location, error in facts['errors']:
        review.add('Error','Spreadsheet formula error',error,[location])
    if facts['uncached']:
        review.add('Needs review','Some calculation results are unavailable',
                   'Save a recalculated copy in Excel and run again. Blank cached results are not treated as zero.',facts['uncached'])
    if not rows or not table:
        review.add('Error','Property inventory is empty','Step 1 and Step 2 need readable property rows.',[facts['step1'],s2])
    eligible=[r for r in table if r['land_use'] in ELIGIBLE]
    invalid=[r for r in table if r['land_use'] not in ELIGIBLE|{'EXC'} or r['area'] is None or r['area']<0]
    for row in invalid:
        review.add('Needs review','Incomplete property inputs','Review the land use and numeric lot area.',[row['location']])
    if len(rows)!=len(table):
        review.compare('Step 1 / Step 2 row count',len(rows),len(table),[facts['step1'],s2])
    linked=[];used=set()
    for r in table:
        matches=[p for p in rows if p['row']==r.get('source_row')] if r.get('source_row') else [p for p in rows if norm(p['description'])==norm(r['description'])]
        if len(matches)==1 and matches[0]['row'] not in used:
            linked.append((matches[0],r));used.add(matches[0]['row'])
        else:
            review.add('Needs review','Step 2 row could not be uniquely linked','Compare this description to Step 1. Area checks never pair rows solely by their position.',[r['location']])
    if linked:
        pairs=[(p,r) for p,r in linked if r['land_use']!='EXC']
        if all(p['area'] is not None and r['area'] is not None for p,r in pairs):
            review.compare('Step 1 / Step 2 eligible property area',sum(p['area'] for p,r in pairs),sum(r['area'] for p,r in pairs),[facts['step1'],s2],tolerance=.1)
        for p,r in linked:
            if r['land_use']!='EXC' and p['area'] is not None and r['area'] is not None and abs(p['area']-r['area'])>.1:
                review.compare('Step 1 / Step 2 lot area',p['area'],r['area'],[p['location'],r['location']],tolerance=.1)
            if r['land_use']=='EXC':
                review.add('Needs review','Excluded lot retained in review','This row remains in the HCAD and map inventory. Confirm its exclusion; a zero calculation area is allowed.',[p['location'],r['location']])
    count=len(eligible); sfr=sum(r['land_use']=='SFR' for r in eligible)
    metrics={'eligible_count':count,'sfr_count':sfr,'sfr_lot_percent':100*sfr/count if count else None,
             'all_count':len(table), 'excluded_count':len(table)-count}
    for cell,expected in [('A13' if facts['renewal'] else 'A15',count),('D9' if facts['renewal'] else 'H11',sfr/count if count else None)]:
        actual=number(facts['cached'][cell])
        if not invalid and actual is not None and expected is not None:
            review.compare('Eligible lot count' if cell.startswith('A') else 'SFR percentage by lot count',expected,actual,[f'{s2}!{cell}',f'{s2} property rows'],tolerance=.0051)
    if not facts['renewal'] and not invalid:
        area=sum(r['area'] for r in eligible); support=sum(r['area'] for r in eligible if r['signed'] in ('Y','YES','TRUE','1'))
        unknown=[r for r in eligible if r['signed'] not in ('','0','N','NO','FALSE','Y','YES','TRUE','1')]
        if unknown:
            review.add('Needs review','Unrecognized support entries','Use the signed application to confirm these entries.',[r['location'] for r in unknown])
        else:
            metrics.update(support_area_percent=100*support/area if area else None, support_lot_percent=100*sum(r['signed'] in ('Y','YES','TRUE','1') for r in eligible)/count if count else None)
            actual=number(facts['cached']['F8'])
            if actual is not None and area:
                review.compare('Support percentage by land area',support/area,actual,[f'{s2}!F8',f'{s2} column C and lot areas'],tolerance=.0051)
        review.add('Needs review','Support signatures','The supplied files do not independently prove who signed. Compare support entries with the signed application/petition.',[f'{s2} signed-in-support column'])
    if minimum is None or minimum<=0:
        review.add('Needs review','Proposed minimum is unavailable','Enter or recalculate the proposed minimum before comparing document values.',[facts['minimum_ref']])
    else:
        developed=[r for r in eligible if r['land_use']!='VAC' and r['measure'] is not None and r['measure']>0] if facts['family']=='mbl' else eligible
        metric_complete=not invalid and all(r['measure'] is not None for r in eligible if facts['family']=='mls' or r['land_use']!='VAC')
        if not metric_complete:
            review.add('Needs review','Qualifying-lot count unavailable','One or more eligible lots lack a readable measurement. Counts were not treated as verified.',[s2+' property rows'])
        if metric_complete:
            metrics['qualifying_count']=sum(r['measure']>=minimum for r in developed)
            metrics['qualifying_denominator']=len(developed)
            metrics['qualifying_area_percent']=100*sum(r['area'] for r in eligible if r['measure'] is not None and r['measure']>=minimum)/sum(r['area'] for r in eligible) if sum(r['area'] for r in eligible) else None
            if facts['renewal'] and not facts['old']:
                unchanged=0
                for r in developed:
                    qualifies=r['measure']>=minimum or (r['original'] is not None and r['measure']==r['original'])
                    unchanged+=qualifies
                    if r['flag'] in ('YES','NO'):
                        review.compare('Renewal unchanged-or-meets flag','YES' if qualifies else 'NO',r['flag'],[f"{s2}!F{r['row']}",r['location']])
                metrics['unchanged_or_meets_count']=unchanged
            if not facts['renewal']:
                m=re.search(r'>=\s*(0\.\d+)',facts['threshold_formula'])
                if m and developed:
                    threshold=float(m[1]); ordered=sorted(developed,key=lambda r:r['measure'],reverse=True)
                    weights=[r['area'] if facts['family']=='mls' else 1 for r in ordered]
                    total=sum(weights); cumulative=0; calculated=None
                    for row,weight in zip(ordered,weights):
                        cumulative+=weight
                        if total and cumulative/total+1e-10>=threshold:
                            calculated=row['measure'];break
                    if calculated is not None:
                        review.compare(f'Proposed minimum under worksheet {threshold:.0%} method',calculated,minimum,[facts['minimum_ref'],f'{s2} property rows'])
                else:
                    review.add('Needs review','Minimum selection method','The worksheet method was not recognized. The reviewer checked stated values but did not approve the minimum-selection rule.',[facts['minimum_ref']])
        if facts['family']=='mbl':
            review.add('Needs review','Building-line measurements and corner lots','HCAD land area does not verify a building setback. Confirm measurements and the planner-entered corner-lot exclusions against the survey/plat.',[s2,'Case Info'])
    if invalid:
        metrics={k:v for k,v in metrics.items() if k in ('all_count','excluded_count')}
    facts['metrics']=metrics
    seen=set()
    for r in rows:
        identity=(r['account'],norm(r['legal']),norm(r['subdivision']))
        if identity in seen:
            review.add('Error','Duplicate account and legal lot row','The same account/legal description occurs more than once. Distinct split lots may share an account.',[r['location']])
        seen.add(identity)
        if not r['account']:
            review.add('Needs review','Missing valid HCAD account','This row cannot be reliably matched to an HCAD record.',[r['location']])


def hcad_checks(review, facts, blocks, rca):
    groups=defaultdict(list)
    for block in blocks:
        record=hcad_record(block['text'])
        if record:
            groups[record['account']].append({**record,'source':block})
    desired=defaultdict(list)
    for row in facts['rows']:
        if row['account']:desired[row['account']].append(row)
    if not groups:
        review.add('Needs review','HCAD records could not be identified','No labelled account numbers were readable. Supply readable HCAD PDFs or enable local OCR.',[])
    for account,rows in desired.items():
        records=groups.get(account,[]); locations=[r['location'] for r in rows]
        if not records:
            review.add('Needs review','HCAD record not found',f'Account {account} has no readable match in the selected HCAD files.',locations);continue
        areas={r['area'] for r in records if r['area'] is not None}
        record=next((r for r in records if r['area'] is not None),records[0]);source=record['source'];locations+=[source['location']]
        review.add('Passed','HCAD account matched',f'Account {account}; {len(rows)} calculation row(s). One HCAD record can support multiple split lots.',locations,source)
        if len(areas)>1:
            review.add('Needs review','Conflicting HCAD versions',f'Account {account} has different land areas in the uploaded records: {sorted(areas)}. Confirm the applicable record year.',locations)
        elif areas and all(r['area'] is not None for r in rows):
            review.compare('HCAD land area / combined lot rows',record['area'],sum(r['area'] for r in rows),locations,source,tolerance=max(.1,record['area']*.00001))
        else:
            review.add('Needs review','Land area comparison incomplete',f'Account {account}: a labelled HCAD land area or calculation area is missing.',locations)
        for key,label in [('address','Property address'),('owner','Owner name'),('legal','Legal description'),('mailing','Mailing address')]:
            supplied={r[key] for r in records if r[key]}
            if not supplied:
                review.add('Needs review',label+' not independently read',f'Account {account}: confirm this field in the HCAD record.',locations);continue
            if len(supplied)>1:
                review.add('Needs review','Conflicting HCAD '+label.lower(),f'Account {account}: uploaded records contain different values.',locations);continue
            actual=next(iter(supplied)); expected={r[key] for r in rows if r[key]}
            if key=='legal' and len(rows)>1:
                from .lot_divisions import parse_lots
                lots=parse_lots(actual)
                if lots:
                    def legal_id(text):
                        m=re.search(r'\bLT\s+(\d+[A-Z]?)\s+BLK\s+([A-Z0-9-]+)\b',text,re.I)
                        return (m[1].upper(),m[2].upper()) if m else None
                    wanted=[legal_id(l['description']) for l in lots];actual_lots=[legal_id(r['legal']) for r in rows]
                    if all(wanted) and all(actual_lots):review.compare('Listed HCAD lots / split rows',sorted(wanted),sorted(actual_lots),locations,source)
                    else:review.add('Needs review','Split-lot legal descriptions','Some individual legal descriptions could not be parsed as whole lots.',locations)
                else:
                    review.add('Needs review','Split-lot legal descriptions',f'Account {account}: confirm partial lots/tracts against the plat. HCAD says {actual}.',locations)
            elif expected:
                keyfn=address_key if key in ('address','mailing') else norm
                agrees=len(expected)==1 and (keyfn(next(iter(expected))) in keyfn(actual) if key=='mailing' else keyfn(next(iter(expected)))==keyfn(actual))
                review.add('Passed' if agrees else 'Needs review',label+' comparison',f'Calculation: {"; ".join(sorted(expected))}. HCAD: {actual}.',locations,source)
        if len(rows)>1:
            review.add('Needs review','Allocated lot areas',f'Account {account}: the sum was checked above. Equal division does not establish surveyed individual lot sizes.',locations)
    for account in sorted(set(groups)-set(desired)):
        review.add('Needs review','HCAD account outside the calculation',f'Uploaded account {account} is not in Step 1.',[r['source']['location'] for r in groups[account]])
    # Match accounts visible in the finished RCA, never merely its attachment filename.
    packet_accounts={r['account'] for b in rca if (r:=hcad_record(b['text']))}
    if desired:
        missing=set(desired)-packet_accounts
        recognized=any(b['method']!='PDF text' for b in rca if hcad_record(b['text']))
        review.add('Needs review' if missing or recognized else 'Passed','HCAD account coverage in RCA',
                   'Not independently located in the RCA: '+', '.join(sorted(missing)) if missing else 'Every Step 1 account was read in the RCA HCAD pages.'+(' Confirm the recognized account numbers against the images.' if recognized else ''),['RCA packet'])


def document_checks(review, facts, blocks, role):
    text='\n'.join(b['text'] for b in blocks); metrics=facts['metrics']; minimum=facts['minimum']
    found_minimum=False; found_count=False
    for b in blocks:
        t=clean(b['text']);loc=[b['location']]
        for m in re.finditer(r'\{\{[^}]+\}\}|\[(?:review|enter|confirm)[^\]]*\]|\bX\s*\(X\)|\b(?:TBD|TODO)\b',t,re.I):
            review.add('Error','Unresolved document placeholder',m[0],loc,b)
        for m in re.finditer(WORD_NUMBER,t,re.I):
            word_value=words_number(m[1]); numeric=number(m[2])
            if word_value is not None and word_value!=numeric:
                review.compare('Written number / numeral',word_value,numeric,loc,b)
        for m in re.finditer(r'\b(MLS|MBL)\s*#?\s*(\d+)\s*(REN)?\b',t,re.I):
            expected=re.sub(r'\s*REN$','',facts['application'],flags=re.I)
            if expected and m[2]!=expected:
                review.compare('Application number',expected,m[2],['Step 2!B2',*loc],b)
        # Check numeric minima in their sentence context, not every parcel area.
        units=r'(?:Square\s*Feet|Square\s*Foot|Sq\.?\s*Ft\.?)' if facts['family']=='mls' else r'(?:Feet|Foot)'
        for m in re.finditer(r'(?:\(([\d,]+(?:\.\d+)?)\)|\b([\d,]+(?:\.\d+)?))\s*[- ]?\s*'+units+r'\b',t,re.I):
            context=t[max(0,m.start()-160):min(len(t),m.end()+100)]
            label='lot size' if facts['family']=='mls' else 'building line'
            if label not in context.lower():continue
            value=number(m[1] or m[2]);found_minimum=True
            if minimum is not None:review.compare('Proposed minimum in '+role,minimum,value,[facts['minimum_ref'],*loc],b,tolerance=.01)
        count=re.search(r'\(([\d,]+)\)\s+out\s+of\s+(?:'+WORD+r'[ -]+)*\(([\d,]+)\)\s+of\s+the\s+(?:lots|developed\s+properties)',t,re.I)
        if count and 'qualifying_count' in metrics:
            found_count=True
            review.compare('Lots meeting proposed minimum',metrics['qualifying_count'],int(count[1].replace(',','')),[facts['step2']+' property rows',*loc],b)
            review.compare('Qualifying-lot denominator',metrics['qualifying_denominator'],int(count[2].replace(',','')),[facts['step2']+' property rows',*loc],b)
        for m in re.finditer(r'(\d+(?:\.\d+)?)\s*%',t):
            before=t[max(0,m.start()-80):m.start()].lower();after=t[m.end():m.end()+130].lower();context=before+after;key=None
            if re.search(r'(?:must be|at least|more than|less than|under|minimum of|threshold)\s*$',before):continue
            if re.search(r'\bsfr\s*:?\s*$',before) or (re.search(r'single.family.*(?:representing|comprise|constitute)\s*$',before) and re.match(r'\s+of\s+(?:the\s+)?(?:eligible\s+)?(?:lots|properties)',after)):
                key='sfr_lot_percent'
            elif re.match(r'\s+of\s+(?:the\s+)?(?:eligible\s+)?(?:application\s+)?(?:land\s+)?area\b',after) and re.search(r'sign|consent|support',context):key='support_area_percent'
            if key and metrics.get(key) is not None:
                review.compare(key.replace('_',' ').capitalize(),metrics[key],float(m[1]),[facts['step2']+' property rows',*loc],b,tolerance=.51)
            elif re.search(r'sign|consent|support',context):
                review.add('Needs review','Support percentage wording',f'The document states {m[1]}% without an unambiguous area denominator. Compare its intended meaning with Step 2 support by area.',loc,b)
        route=re.search(r'Total\s+Lots\s*/\s*Lots\s+That\s+Meet\s+(?:Building\s+Line|Lot\s+Size)\s*:\s*(\d+)\s*/\s*(\d+)',t,re.I)
        if route:
            metric='unchanged_or_meets_count' if facts['renewal'] else 'qualifying_count'
            if metric in metrics:
                review.compare('Routing qualifying lot count',metrics[metric],int(route[2]),[facts['step2']+' property rows',*loc],b)
                review.compare('Routing eligible lot count',metrics.get('eligible_count'),int(route[1]),[facts['step2']+' property rows',*loc],b)
        district=re.search(r'Council\s+District\s*:?\s*([A-K])\b',t,re.I)
        expected_district=clean(facts['case_info'].get('Council_District')).upper()
        if district and expected_district:
            review.compare('Council district',expected_district,district[1].upper(),['Case Info!Council_District',*loc],b)
        noncorner=re.search(r'\b(\d+)\s+(?:lots|properties)\s*\(?excluding\s+corner\s+lots',t,re.I)
        if not noncorner:
            noncorner=re.search(r'\((\d+)\)\s+(?:lots|properties)\s*\(?excluding\s+corner\s+lots',t,re.I)
        entered=number(facts['case_info'].get('PC_Noncorner_Lot_Count'))
        if noncorner and entered is not None:
            review.compare('Planner-entered noncorner lot count',entered,int(noncorner[1]),['Case Info!PC_Noncorner_Lot_Count',*loc],b)
        expected_range=re.search(r'\b(\d{2,5})(?:\s*(?:-|–|and|through|to)\s*(\d{2,5}))?',facts['block'],re.I)
        for boundary in re.finditer(r'\b(\d{2,5})(?:\s*(?:-|–|and|through|to)\s*(\d{2,5}))?\s+blocks?\s+(?:of\s+)?([^,;\n]{2,100})',t,re.I):
            if expected_range and address_key(boundary[3]).startswith(address_key(facts['street'])):
                expected='-'.join(v for v in expected_range.groups() if v);actual='-'.join(v for v in boundary.groups()[:2] if v)
                review.compare('Application block range',expected,actual,['Step 2!D4',*loc],b)
    if not found_minimum:
        review.add('Needs review','Proposed minimum not located in '+role,'No unambiguous minimum value was readable in the supplied document.',[role])
    if role=='Draft ordinance' and not facts['renewal'] and not found_count:
        review.add('Needs review','Qualifying-lot statement not located','The reviewer could not locate the written-number/numeral lot-count statement.',[role])
    # Boundary words can occur across wrapped lines. Missing items are review
    # notes because abbreviations and exhibits can change presentation.
    keyed=address_key(text)
    for label,expected in [('Application street',facts['street']),('First cross street',facts['cross'][0]),('Second cross street',facts['cross'][1])]:
        if expected and address_key(expected) not in keyed:
            review.add('Needs review',label+' not located in '+role,f'Expected {expected}. Check the description and map exhibits.',[role,'Step 2 boundary fields'])
    if facts['renewal']:
        from .designation_lookup import ordinance, day
        ci=facts['case_info'];expected=ordinance(ci.get('Original_Ordinance_No'))
        refs=[]
        for b in blocks:
            for m in re.finditer(r'Ordinance\s+No\.?\s*(\d{4}\s*-\s*\d+)\s+(?:effective|expired|established)',clean(b['text']),re.I):
                refs.append((ordinance(m[1]),b))
        if expected and refs:
            for value,b in refs:review.compare('Previous ordinance reference',expected,value,['Case Info!Original_Ordinance_No',b['location']],b)
        else:review.add('Needs review','Previous ordinance reference incomplete','Confirm the original ordinance number in Case Info and the finished document.',[role,'Case Info'])
        for field,phrase in [('Original_Effective_Date',r'effective\s+(?:on\s+)?'),('Original_Expiration_Date',r'expired\s+(?:on\s+)?')]:
            expected=day(ci.get(field));found=False
            for b in blocks:
                for m in re.finditer(phrase+r'([A-Z][a-z]+\s+\d{1,2},\s*\d{4}|\d{1,2}/\d{1,2}/\d{4})',clean(b['text'])):
                    found=True;actual=day(m[1])
                    if expected and actual:review.compare(field.replace('_',' '),str(expected),str(actual),['Case Info!'+field,b['location']],b)
            if not found or not expected:review.add('Needs review',field.replace('_',' ')+' comparison incomplete','Confirm against the original ordinance; an omitted date is not assumed.',[role,'Case Info!'+field])
    if role=='Draft ordinance':
        for letter in 'ABC':
            if not re.search(r'EXHIBIT\W{0,6}'+letter+r'\b',text,re.I):review.add('Needs review','Ordinance exhibit '+letter+' not located','Check that the cited exhibit is attached and complete.',[role])
        review.add('Needs review','Draft signature and adoption blanks','Adoption dates, signatures, and the new ordinance number can legitimately remain blank in a draft. They are not automatically treated as errors.',[role])


def map_checks(review, facts, blocks):
    maps=[]
    for b in blocks:
        text=b['text']
        if re.search(r'Special\s+Minimum\s+(?:Lot\s+Size|Building\s+Line)',text,re.I) and re.search(r'Appraisal\s+District|\bFeet\b.*\bSource\b|Boundary\s*$',text,re.I|re.M):maps.append(b)
    if not maps:
        review.add('Needs review','Boundary map not identified in RCA','No readable map legend was identified. Open the RCA and confirm the boundary and tax maps are attached.',['RCA packet']);return
    numbers={re.match(r'\d+',r['address'])[0] for r in facts['rows'] if re.match(r'\d+',r['address'])}
    numbers.update({'0'} if any(not re.match(r'[1-9]\d*\s',r['address']) for r in facts['rows']) else set())
    for b in maps:
        review.add('Passed','Map legend located','Map text was found inside the uploaded RCA.',[b['location']],b)
        if re.search(r'Harris\s+County\s+Appraisal\s+District',b['text'],re.I):
            review.add('Error','Outdated map source name','Use Harris Central Appraisal District.',[b['location']],b)
        missing=sorted(numbers-set(re.findall(r'\b\d+\b',b['text'])))
        if missing:review.add('Needs review','Map address labels not located','These labels were not independently read: '+', '.join(missing)+'. Check the map image; small/rotated labels can be missed.',[b['location']])
        for street in [facts['street'],*facts['cross']]:
            if street and address_key(street) not in address_key(b['text']):review.add('Needs review','Map street name not located',street,[b['location']])
        review.add('Needs review','Boundary and address placement','Confirm all intended parcels, including excluded properties, are enclosed; the boundary is contiguous; labels sit in their own lots; and interior lot lines agree with the plat. A flattened RCA image does not provide enough geographic evidence to certify these checks.',[b['location']])
    if not any(re.search(r'HCAD[_ ]Map|Account\s+Detail.*Map|Facet\s+Page',b['text'],re.I|re.S) for b in blocks):
        review.add('Needs review','Tax map not identified','Confirm the HCAD tax map and its yellow boundary are in the RCA.',['RCA packet'])


def run_review(inputs, wizard, progress=None, cancel=None, use_ocr=True):
    """Snapshot inputs, compare independently, and fingerprint evidence for reruns."""
    review=Review(wizard);progress=progress or (lambda t:None)
    entries=[('Calculation sheet',inputs.get('calculation','')),
             ('RCA',inputs.get('rca','')),('Draft ordinance',inputs.get('ordinance',''))]
    hcad=inputs.get('hcad',[])
    if isinstance(hcad,str):hcad=[p for p in hcad.split('||') if p]
    entries.extend(('HCAD',p) for p in dict.fromkeys(hcad))
    if not hcad:review.add('Needs review','HCAD files not selected','Select the HCAD records to compare property facts.',['HCAD'])
    docs=defaultdict(list);facts=None
    with tempfile.TemporaryDirectory(prefix='rca_peer_review_') as temp:
        for index,(role,path) in enumerate(entries):
            check_cancel(cancel)
            if not path:
                review.add('Needs review',role+' not selected','Select the completed file and run again.',[role]);continue
            path=Path(path)
            try:
                progress('Reading '+path.name+'...');raw=read_shared_bytes(path)
                if len(raw)>250_000_000:raise ValueError('File exceeds the 250 MB review limit.')
                signature=hashlib.sha256(raw).hexdigest()
                snapshot=Path(temp)/f'{index}{path.suffix}';snapshot.write_bytes(raw)
                review.result['files'].append({'role':role,'path':str(path.resolve()),'sha256':signature,'bytes':len(raw)})
                if role=='Calculation sheet':
                    facts=read_calculation(snapshot)
                    if facts['wizard']!=wizard:
                        review.add('Needs review','Workbook belongs to another wizard',f"Selected tab: {wizard}; workbook: {facts['wizard']}. Review uses the workbook's type.",['Step 2!A1'])
                    calculation_checks(review,facts)
                else:
                    blocks=read_document(snapshot,lambda t:progress(t.replace(snapshot.name,path.name)),cancel,use_ocr)
                    for b in blocks:
                        b['location']=f'{path.name}, {b["location"]}';b['path']=str(path.resolve())
                        if b.get('warning'):review.add('Needs review','Document reading incomplete',b['warning'],[b['location']])
                        if b['method']=='OCR':review.add('Needs review','Recognized text needs confirmation','This page/image was read with local OCR. Small text and handwriting can be misread.',[b['location']])
                    if not any(clean(b['text']) for b in blocks):review.add('Needs review','No readable text in '+role,'The file was not checked successfully. Use a readable copy or enable OCR.',[str(path)])
                    docs[role].extend(blocks)
            except InterruptedError:raise
            except Exception as error:
                review.add('Needs review','Could not read '+role,str(error),[str(path)])
        if facts:
            hcad_checks(review,facts,docs['HCAD'],docs['RCA'])
            for role in ('RCA','Draft ordinance'):
                if docs[role]:document_checks(review,facts,docs[role],role)
            map_checks(review,facts,docs['RCA'])
            review.result['facts']={k:v for k,v in facts.items() if k not in ('rows','table','errors','cached','threshold_formula')}
        else:review.add('Needs review','Cross-document checks unavailable','A readable calculation sheet is required. No consistency approval was issued.',[])
        for item in review.result['files']:
            check_cancel(cancel)
            try:changed=hashlib.sha256(read_shared_bytes(item['path'])).hexdigest()!=item['sha256']
            except OSError:changed=True
            if changed:
                review.result['status']='stale'
                review.add('Needs review','Source file changed during review','Run again to check the newest saved version.',[item['path']])
    review.result['counts']=dict(Counter(f['severity'] for f in review.result['findings']))
    return review.result


def report_html(result):
    escape=lambda v:html.escape(str(v))
    rows=[]
    for severity in ('Error','Needs review','Passed'):
        for f in result['findings']:
            if f['severity']!=severity:continue
            rows.append(f'<tr><td class="{severity.split()[0].lower()}">{escape(severity)}</td><td><b>{escape(f["title"])}</b><p>{escape(f["details"])}</p></td><td>{"<br>".join(escape(x) for x in f["locations"])}</td></tr>')
    files=''.join(f'<li>{escape(f["role"])}: {escape(f["path"])}<br><small>SHA-256 {f["sha256"]}</small></li>' for f in result['files'])
    return '<!doctype html><html lang="en"><meta charset="utf-8"><title>Case Peer Review</title><style>body{font:15px Segoe UI,Arial,sans-serif;color:#18263a;max-width:1120px;margin:36px auto;padding:0 20px}h1{margin-bottom:8px}table{border-collapse:collapse;width:100%}th,td{text-align:left;vertical-align:top;border-bottom:1px solid #ccd4df;padding:12px}th{background:#eaf0f6}td:first-child{width:100px}td:last-child{width:28%;overflow-wrap:anywhere}p{margin:8px 0}small{overflow-wrap:anywhere}.error{color:#a52424}.needs{color:#8a5700}.passed{color:#236244}@media print{body{margin:0;font-size:10pt}tr{break-inside:avoid}thead{display:table-header-group}}</style><h1>Case Peer Review</h1><p>'+escape(result['wizard'].replace('_',' ').upper())+' · '+escape(result['created_utc'])+'</p><p>'+escape(' | '.join(f'{key}: {result["counts"].get(key,0)}' for key in ('Error','Needs review','Passed')))+'</p><p>Advisory review of the saved files listed below. Needs review items are unresolved, not passed. Sources were not modified.</p><table><thead><tr><th>Result</th><th>Finding and evidence</th><th>Location</th></tr></thead><tbody>'+''.join(rows)+'</tbody></table><h2>Reviewed files</h2><ul>'+files+'</ul></html>'


def save_report(result,path):
    path=Path(path)
    # A single self-contained, printable HTML report avoids companion-file
    # overwrites and can be opened without Word, Excel, or an internet connection.
    if path.resolve() in {Path(f['path']).resolve() for f in result['files']}:
        raise ValueError('Choose a report filename different from the source files.')
    raw=report_html(result).encode('utf8')
    import os
    fd,temp=tempfile.mkstemp(prefix='.peer_report_',suffix='.tmp',dir=path.parent)
    try:
        with os.fdopen(fd,'wb') as stream:stream.write(raw)
        os.replace(temp,path)
    finally:
        if Path(temp).exists():Path(temp).unlink()
    return str(path)
