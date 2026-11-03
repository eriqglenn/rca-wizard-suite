"""Local PDF reading and conservative suggestions for the City's application forms.

OCR output is untrusted source text, never code or instructions. Every suggested
field remains editable and requires planner review before creating a workbook.
"""
from __future__ import annotations

from datetime import datetime
import json
import os
from pathlib import Path
import re
import subprocess
import tempfile

import fitz

from .paths import WIZARD_FILES_DIR


def normalize(text):
    return re.sub(r'\s+', ' ', str(text or '')).strip()


def parse_boundary(text):
    text = normalize(text).strip(' ,.;')
    out = {'boundary_description': text}
    # Do not use example text printed on the blank application form.
    if not text or re.search(r'Golden Retriever|Cocker Spaniel|Schnauzer|^Example\s*:', text, re.I):
        return {}
    block = re.search(r'\b(\d{2,5}(?:\s*(?:-|–|and|to|through)\s*\d{2,5})?)\s+blocks?\s+(?:of\s+)?', text, re.I)
    if block:
        out['block_range'] = re.sub(r'\s*(?:and|to|through|–|-)\s*', '-', block[1], flags=re.I)
        rest = text[block.end():]
        street = re.split(r'[,;]|\b(?:north|south|east|west|both|between)\b', rest, maxsplit=1, flags=re.I)[0]
        out['street'] = normalize(street).strip(' .')
    between = re.search(r'\bbetween\s+(.+?)\s+and\s+(.+)$', text, re.I)
    if between:
        out['cross_1'], out['cross_2'] = [normalize(x).strip(' ,.;') for x in between.groups()]
    # Compass words in cross-street names are not sides of the application.
    before = text[:between.start()] if between else text
    sides = re.search(r'\b(north|south|east|west)(?:\s+and\s+(north|south|east|west))?\s+sides?\b', before, re.I)
    if sides:
        out['sides'] = ' and '.join(x.lower() for x in sides.groups() if x)
    if not out.get('street'):
        match = re.search(r'\bsides?\s+of\s+(.+?)(?:\s+between\b|$)', text, re.I)
        if match:
            out['street'] = normalize(match[1]).strip(' ,.;')
    return out


def normalized_words(page, rotation=0):
    width, height = page.rect.width, page.rect.height
    return [{'text': w[4], 'x': w[0]/width, 'y': w[1]/height,
             'w': (w[2]-w[0])/width, 'h': (w[3]-w[1])/height} for w in page.get_text('words')]


def lines_from_words(words):
    rows = []
    for word in sorted(words, key=lambda w: (w['y'] + w['h']/2, w['x'])):
        cy = word['y'] + word['h']/2
        row = next((r for r in reversed(rows[-5:]) if abs(r['y']-cy) < max(.004, word['h']*.45)), None)
        if row is None:
            row = {'y': cy, 'words': []}
            rows.append(row)
        row['words'].append(word)
    return [{'text': ' '.join(w['text'] for w in sorted(r['words'], key=lambda w: w['x'])),
             'y': r['y'], 'words': sorted(r['words'], key=lambda w: w['x'])} for r in rows]


def extract_cover(words):
    lines = lines_from_words(words)
    if any(re.search(r'Applicant.{0,4}\s+Name\s*:',r['text'],re.I) for r in lines):
        return extract_current_cover(words)
    out = {}
    def anchor(pattern, low=0, high=1):
        return next((r for r in lines if low < r['y'] < high and re.search(pattern, r['text'], re.I)), None)
    def region(x0, x1, y0, y1):
        return normalize(' '.join(r['text'] for r in lines_from_words([
            w for w in words if x0 <= w['x']+w['w']/2 <= x1 and y0 <= w['y']+w['h']/2 <= y1])))
    name = anchor(r'Full\s+Name', 0, .5)
    address = anchor(r'^Address', 0, .5)
    phone = anchor(r'^Phone', 0, .5)
    if name:
        out['applicant_name'] = region(.165, .70, name['y']-.013, name['y']+.013)
        date = region(.72, 1, name['y']-.02, name['y']+.02)
        match = re.search(r'\b\d{1,2}/\d{1,2}/\d{2,4}\b', date)
        if match: out['application_date'] = match[0]
    if address:
        out['applicant_address'] = region(.16, .74, address['y']-.015, address['y']+.012)
    city = anchor(r'\bCity\b.*\bState\b.*ZIP', .20, .45)
    if city:
        out['applicant_city'] = region(.16, .62, city['y']-.041, city['y']-.012)
        out['applicant_state'] = region(.63, .75, city['y']-.041, city['y']-.012)
        out['applicant_zip'] = region(.75, 1, city['y']-.041, city['y']-.012)
    if phone:
        out['applicant_phone'] = region(.16, .48, phone['y']-.014, phone['y']+.014)
        out['applicant_email'] = re.sub(r'^Email\s*:?\s*', '', region(.48, 1, phone['y']-.014, phone['y']+.014), flags=re.I)
    desc = anchor(r'Describe.*location.*application', .3, .7)
    if desc:
        sign = anchor(r'Signature', desc['y'], .7)
        boundary = region(0, 1, desc['y']+.029, (sign['y']-.012) if sign else desc['y']+.09)
        out.update(parse_boundary(boundary))
    file_no = anchor(r'File\s+Number', .75)
    if file_no:
        value = region(.22, .39, file_no['y']-.016, file_no['y']+.008)
        if re.fullmatch(r'\d{1,5}', value): out['application_number'] = value
    district = anchor(r'City\s+Council\s+District', .75)
    if district:
        value = region(.26, .39, district['y']-.007, district['y']+.01)
        if re.fullmatch(r'[A-K]', value, re.I): out['council_district'] = value.upper()
        out['planner_name'] = region(.59, 1, district['y']-.01, district['y']+.01)
    # Received stamp is independent of the application/signature date.
    for line in lines:
        if line['y'] > .18: continue
        match = re.search(r'\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s*(\d{1,2})\s*(20\d{2})\b', line['text'], re.I)
        if match:
            try: out['date_received'] = datetime.strptime(' '.join(match.groups()), '%b %d %Y').strftime('%m/%d/%Y')
            except ValueError: pass
    # Drop obvious labels accidentally included when handwriting is unreadable.
    for key, val in list(out.items()):
        if re.search(r'Apartment|Street Address|ZIP Code|Signature|Planner Assigned|STAFF USE', val, re.I): out[key] = ''
    if not re.match(r'^\d+\s+[A-Za-z]',out.get('applicant_address','')): out.pop('applicant_address',None)
    if not re.fullmatch(r'\d{5}(?:-\d{4})?',out.get('applicant_zip','')): out.pop('applicant_zip',None)
    if not re.fullmatch(r'[A-Z]{2}',out.get('applicant_state','')): out.pop('applicant_state',None)
    if len(re.sub(r'\D','',out.get('applicant_phone',''))) not in (10,11): out.pop('applicant_phone',None)
    if not re.fullmatch(r'[^\s@]+@[^\s@]+\.[^\s@]+',out.get('applicant_email','')): out.pop('applicant_email',None)
    return {k: v for k, v in out.items() if v}


def extract_current_cover(words):
    """Read the 2026 forms, whose boundary precedes the main applicant section."""
    lines=lines_from_words(words);out={}
    def anchor(pattern,low=0,high=1):
        return next((r for r in lines if low<r['y']<high and re.search(pattern,r['text'],re.I)),None)
    def region(x0,x1,y0,y1):
        selected=[w for w in words if x0<=w['x']+w['w']/2<=x1 and y0<=w['y']+w['h']/2<=y1 and not re.fullmatch(r'[_\-.]+',w['text'])]
        return normalize(' '.join(r['text'] for r in lines_from_words(selected))).strip(' _:')
    main=anchor(r'\bMain\s*Applicant');co=anchor(r'\bCo\s*-?\s*Applicant')
    if main:
        top=main['y'];bottom=co['y'] if co else min(.8,top+.28)
        name=anchor(r'\bName\s*:',top,bottom)
        phone=anchor(r'Phone\s+Number',top,bottom)
        address=anchor(r'^Address\s*:',top,bottom)
        signature=anchor(r'^Signature',top,bottom)
        if name:out['applicant_name']=region(.21,.97,name['y']-.016,name['y']+.012)
        if phone:
            number=region(.19,.43,phone['y']-.016,phone['y']+.012)
            if len(re.sub(r'\D','',number)) in (10,11):out['applicant_phone']=number
            email=region(.56,.98,phone['y']-.016,phone['y']+.012)
            if re.fullmatch(r'[^\s@]+@[^\s@]+\.[^\s@]+',email):out['applicant_email']=email
        if address:
            value=region(.137,.97,address['y']-.016,address['y']+.012)
            if re.match(r'^(?:\d+\s+|P\.?\s*O\.?\s+Box\b)',value,re.I):
                match=re.fullmatch(r'(.+?),\s*([^,]+?),?\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)',value,re.I)
                if not match:match=re.fullmatch(r'(.+?)\s+(Houston)\s+(TX)\s+(\d{5}(?:-\d{4})?)',value,re.I)
                if match:
                    out.update(zip(('applicant_address','applicant_city','applicant_state','applicant_zip'),match.groups()))
                    out['applicant_state']=out['applicant_state'].upper()
                else:out['applicant_address']=value
        if signature:
            value=region(.76,.98,signature['y']-.018,signature['y']+.012)
            match=re.search(r'\b\d{1,2}/\d{1,2}/\d{2,4}\b',value)
            if match:out['application_date']=match[0]
    desc=anchor(r'Describe.*location.*application');meeting=anchor(r'Date\s+and\s+name.*Planner')
    if desc:
        end=meeting['y']-.005 if meeting else desc['y']+.10
        rows=[r for r in lines if desc['y']+.012<r['y']<end and not re.match(r'\s*Example\b',r['text'],re.I)]
        out.update(parse_boundary(' '.join(r['text'] for r in rows)))
    file_no=anchor(r'^File\s+Number',.7)
    if file_no:
        value=region(.18,.39,file_no['y']-.016,file_no['y']+.012)
        if re.fullmatch(r'\d{1,5}(?:\s*REN)?',value,re.I):out['application_number']=value
    district=anchor(r'^Council\s+District',.7)
    if district:
        value=region(.18,.39,district['y']-.016,district['y']+.012)
        if re.fullmatch(r'[A-K]',value,re.I):out['council_district']=value.upper()
        out['planner_name']=region(.69,.98,district['y']-.016,district['y']+.012)
    for line in lines:
        if line['y']>.18:continue
        match=re.search(r'\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s*(\d{1,2})\s*(20\d{2})\b',line['text'],re.I)
        if match:
            try:out['date_received']=datetime.strptime(' '.join(match.groups()),'%b %d %Y').strftime('%m/%d/%Y')
            except ValueError:pass
    return {k:v for k,v in out.items() if v}


def table_hints(page):
    """Read printed roster rows only; signature pages are not a parcel inventory."""
    lines = lines_from_words(page['words'])
    header = next((r for r in lines if re.search(r'\bTax\s*ID|\bAccount', r['text'], re.I)), None)
    if not header or not re.search(r'LAND\s*SF|Lot\s*Size|OWNER|Street', page['text'], re.I):
        return []
    # Locate columns in the header (often two lines) and preserve row positions.
    near = [w for w in page['words'] if abs(w['y']+w['h']/2-header['y']) < .045]
    land = next((w for w in near if re.match(r'^(?:LAN[DL]|Lot\b)', w['text'], re.I)), None)
    street = next((w for w in near if re.fullmatch(r'Street', w['text'], re.I)), None)
    # Rosters with ADDRESS followed by STREET use the Address column for numbers.
    address = next((w for w in near if re.fullmatch(r'ADDRESS', w['text'], re.I)), None)
    number_x = address['x'] if address else (street['x'] if street else None)
    account_words = [w for w in page['words'] if w['y'] > header['y'] and re.fullmatch(r'\d{11,13}', w['text'])]
    hints = []
    for word in account_words:
        cy = word['y']+word['h']/2
        row = [w for w in page['words'] if abs(w['y']+w['h']/2-cy) < max(.004,word['h']*.5)]
        number = next((w['text'] for w in row if number_x is not None and number_x-.01 <= w['x'] <= number_x+.085 and re.fullmatch(r'\d{2,5}', w['text'])), '')
        area = next((w['text'].replace(',','') for w in row if land and land['x']-.01 <= w['x'] <= land['x']+.085 and re.fullmatch(r'\d[\d,]*(?:\.\d+)?', w['text'])), '')
        hints.append({'account': word['text'].zfill(13), 'number': number,
                      'area': float(area) if area else None, 'page': page['page'],
                      'source_text': ' '.join(w['text'] for w in sorted(row,key=lambda w:w['x']))})
    return hints


def detected_wizard(text):
    """Identify the form heading, without letting it switch the user's wizard."""
    text=normalize(text)
    if re.search(r'Special\s+Minimum\s+Building\s+Line',text,re.I): family='mbl'
    elif re.search(r'Special\s+Minimum\s+Lot\s+Size',text,re.I): family='mls'
    else: return None
    header=text[:500]
    renewal=bool(re.search(r'Application\s+to\s+Renew\b|Renewal\s+Application|Block\s+Renewal\b',text,re.I)
                 or (re.search(r'Application\s+to\b',header,re.I) and re.search(r'\bRenew\b',header,re.I)))
    return family+('_renewal' if renewal else '')


def designation_suggestions(pages):
    """Only labelled minima or an attached boundary-map caption are suggestions.

    Never use a lot-area roster, a handbook example or a parcel's HCAD area as
    the existing minimum. Conflicting captions remain blank for review.
    """
    candidates={'minimum_lot_size':set(),'minimum_building_line':set()}
    sources=[]
    for page in pages:
        text=page['text']
        if re.search(r'Golden Retriever|Cocker Spaniel|Schnauzer',text,re.I): continue
        for key,label,units in [('minimum_lot_size',r'Lot\s+Size',r'Square\s+Feet|Sq\.?\s*Ft\.?'),
                                ('minimum_building_line',r'Building\s+Line',r'Feet|Foot')]:
            pattern=rf'(?:{label}\s+to\s+Renew|Existing\s+Minimum\s+{label})(?:\s*\(in\s+(?:{units})\))?\s*:?\s*([\d,]+(?:\.\d+)?)\b'
            values=re.findall(pattern,text,re.I)
            # A map caption explicitly identifies both the designation and units.
            if re.search(rf'Special\s+Minimum\s+{label}',text,re.I) and re.search(r'\bBoundary\b',text,re.I) and re.search(r'\bReference\s*:',text,re.I):
                values+=re.findall(rf'\b([\d,]+(?:\.\d+)?)\s+(?:{units})\b',text,re.I)
            for value in values:
                value=value.replace(',','')
                if float(value)>0:
                    candidates[key].add(value)
                    sources.append({'field':key,'value':value,'page':page['page']})
    fields={key:next(iter(values)) for key,values in candidates.items() if len(values)==1}
    warnings=[f'Multiple {key.replace("_"," ")} values appear in the PDF. Confirm the existing designation; the suggested field was left blank.' for key,values in candidates.items() if len(values)>1]
    return fields,sources,warnings


def read_small_account_text(path, pages, cancel=None):
    """Re-read the labelled account region in tiny HCAD printouts at 6x.

    Cropping keeps Windows OCR within its image-size limit and avoids guessing
    account digits from a filename or an address shared by several parcels.
    """
    script = WIZARD_FILES_DIR / 'assets' / 'application_import' / 'windows_ocr.ps1'
    powershell = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'System32/WindowsPowerShell/v1.0/powershell.exe'
    if not script.exists() or not powershell.exists(): return ''
    with fitz.open(path) as doc, tempfile.TemporaryDirectory(prefix='rca_account_') as temp:
        for data in pages:
            if cancel and cancel.is_set(): raise ValueError('Record reading cancelled.')
            word=next((w for w in data['words'] if re.fullmatch(r'Account[:#]?',w['text'],re.I)),None)
            if not word: continue
            page=doc[data['page']-1];rect=page.rect
            crop=fitz.Rect(max(0,word['x']-.02)*rect.width,max(0,word['y']-.025)*rect.height,
                          min(1,word['x']+.48)*rect.width,min(1,word['y']+.055)*rect.height)
            scale=min(6,3500/max(crop.width,crop.height))
            png=Path(temp)/'account.png'
            page.get_pixmap(matrix=fitz.Matrix(scale,scale),clip=crop,alpha=False).save(png)
            result=subprocess.run([str(powershell),'-NoProfile','-NonInteractive','-File',str(script),'-ImagePath',str(png)],
                                  capture_output=True,timeout=50,creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
            output=Path(str(png)+'.ocr.json')
            if result.returncode or not output.exists(): continue
            raw=json.loads(output.read_text(encoding='utf-8-sig'))
            text='\n'.join(line['text'] for line in raw['lines'])
            if re.search(r'Account\s*[:#]?\s*\d{13}\b',text,re.I): return text
    return ''


def read_application(path, progress=None, cancel=None, *, force_ocr=False, page_limit=None):
    progress = progress or (lambda s: None)
    path = Path(path).resolve()
    if path.suffix.lower() != '.pdf': raise ValueError('Select an application PDF.')
    pages, warnings = [], []
    script = WIZARD_FILES_DIR / 'assets' / 'application_import' / 'windows_ocr.ps1'
    powershell = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'System32/WindowsPowerShell/v1.0/powershell.exe'
    with fitz.open(path) as doc, tempfile.TemporaryDirectory(prefix='rca_application_') as temp:
        if doc.needs_pass: raise ValueError('The PDF is password protected. Select an unlocked copy.')
        if len(doc) > 80: raise ValueError('Use an application packet with 80 pages or fewer.')
        for i, page in enumerate(doc):
            if page_limit is not None and i >= page_limit: break
            if cancel and cancel.is_set(): raise ValueError('Application reading cancelled.')
            progress(f'Reading application page {i+1} of {len(doc)}...')
            words = normalized_words(page)
            method = 'PDF text'
            native_text = ' '.join(w['text'] for w in words)
            broken_text = sum(ord(c) < 32 or c == '\ufffd' for c in native_text) > max(5, len(native_text)*.05)
            if force_ocr or len(words) < 15 or broken_text:
                method = 'Windows OCR'
                words = []
                if not powershell.exists() or not script.exists():
                    warnings.append(f'Page {i+1}: local text recognition is unavailable. Read the scan in the preview and enter the fields.')
                else:
                    best = None
                    rotations = [0, 90, 270]
                    for rotation in rotations:
                        png = Path(temp)/f'page_{i}_{rotation}.png'
                        scale = min(3, 3500/max(page.rect.width,page.rect.height))
                        pix = page.get_pixmap(matrix=fitz.Matrix(scale,scale).prerotate(rotation), alpha=False)
                        pix.save(png)
                        result = subprocess.run([str(powershell), '-NoProfile', '-NonInteractive', '-File', str(script), '-ImagePath', str(png)],
                                                capture_output=True, timeout=50, creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
                        json_file = Path(str(png)+'.ocr.json')
                        if result.returncode or not json_file.exists():
                            warnings.append(f'Page {i+1}: text recognition could not read this scan; use the page preview.')
                            break
                        raw = json.loads(json_file.read_text(encoding='utf-8-sig'))
                        candidate = [{'text': w['text'], 'x': w['x']/pix.width, 'y': w['y']/pix.height,
                                      'w':w['width']/pix.width, 'h':w['height']/pix.height} for line in raw['lines'] for w in line['words']]
                        text = ' '.join(w['text'] for w in candidate)
                        score = sum(len(re.findall(pattern,text,re.I)) for pattern in [r'\bStreet\b',r'\bOwner\b',r'\bLot\b',r'\bBuilding\b',r'\bAccount\b',r'\bResidential\b'])
                        if best is None or score > best[0]: best = score,candidate,rotation
                        if rotation == 0 and score >= 3: break
                    if best: words = best[1]
            lines = lines_from_words(words)
            pages.append({'page':i+1, 'method':method, 'words':words, 'text':'\n'.join(r['text'] for r in lines)})
    cover = next((p for p in pages if re.search(r'Main.*Applicant',p['text'],re.I) and re.search(r'Minimum.*(?:Lot|Building)',p['text'],re.I)), pages[0] if pages else {})
    fields = extract_cover(cover.get('words',[]))
    minimums,minimum_sources,minimum_warnings=designation_suggestions(pages)
    fields.update(minimums);warnings.extend(minimum_warnings)
    warnings.insert(0, 'Check every suggested field against the application. Handwriting, crossed-out entries and received stamps may be missed or misread.')
    hints = [hint for p in pages for hint in table_hints(p)]
    return {'schema':2,'source_pdf':str(path),'fields':fields,'pages':pages,'cover_page':cover.get('page',1),
            'detected_wizard':detected_wizard(cover.get('text','')),'designation_sources':minimum_sources,
            'property_hints':hints,'warnings':warnings}
