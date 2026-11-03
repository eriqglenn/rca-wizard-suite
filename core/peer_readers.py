"""Read saved review evidence locally. Never recalculate or edit source files."""
from __future__ import annotations
import io
import json
import math
import os
import re
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import fitz
from openpyxl import load_workbook
from .hcad_identity import clean_account
from .workbooks import read_shared_bytes


def clean(value):
    return re.sub(r'\s+', ' ', str(value if value is not None else '')).strip()


def number(value):
    try:
        result = float(str(value).replace(',', '').strip())
        return result if math.isfinite(result) else None
    except (ValueError, TypeError):
        return None


def norm(value):
    return re.sub(r'[^A-Z0-9]', '', clean(value).upper())


def check_cancel(cancel):
    if cancel and cancel.is_set():
        raise InterruptedError('Peer review cancelled.')


def read_calculation(path):
    raw = read_shared_bytes(path)
    book = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    formulas = load_workbook(io.BytesIO(raw), read_only=True, data_only=False)
    try:
        s1 = next(s for s in book if s.title.lower().startswith('step 1'))
        s2 = next(s for s in book if s.title.lower().startswith('step 2'))
        title = clean(s2['A1'].value).upper()
        family = 'mbl' if 'BUILDING LINE' in title else 'mls' if 'LOT SIZE' in title else ''
        if not family:
            raise ValueError('Step 2 does not identify an MLS or MBL calculation sheet.')
        renewal = 'RENEWAL' in title
        wizard = family + ('_renewal' if renewal else '')
        values = {}; expressions = {}
        for s in (s1,s2):
            values[s.title] = {c.coordinate:c.value for row in s.iter_rows(max_row=min(s.max_row,5000),max_col=min(s.max_column,65)) for c in row if c.value is not None}
            fs=formulas[s.title]
            expressions[s.title] = {c.coordinate:c.value for row in fs.iter_rows(max_row=min(fs.max_row,5000),max_col=min(fs.max_column,65)) for c in row if c.value is not None}
        # Resolve literal cells and direct worksheet links independently of the
        # document generator. Other uncached formulas stay unknown, not zero.
        def value(sheet, ref, depth=0):
            val = values.get(sheet,{}).get(ref)
            formula = expressions.get(sheet,{}).get(ref)
            if isinstance(formula, str) and formula.startswith('='):
                direct = re.fullmatch(r"='([^']+)'!\$?([A-Z]+)\$?(\d+)", formula)
                if direct and depth < 3 and direct[1] in book:
                    return value(direct[1], direct[2] + direct[3], depth + 1)
            return val
        heads = {norm(c.value): c.column for c in s1[1] if c.value}
        rows = []
        for rn, rawrow in enumerate(s1.iter_rows(min_row=2, max_row=min(s1.max_row, 5000), values_only=True), 2):
            def field(*names):
                return next((rawrow[heads[norm(n)] - 1] for n in names if norm(n) in heads and rawrow[heads[norm(n)] - 1] not in (None, '')), None)
            account = next((a for name in ('TAX_ID','ALT_TAX_ID','HCAD_NUM','HCAD_NUM_1') if (a:=clean_account(field(name)))), '')
            address = clean(field('Site_addr_1', 'site_address', 'address'))
            legal = clean(field('Legal_Dscr_1', 'Legal Description'))
            if not any((account, address, legal)):
                continue
            if not address:
                address = ' '.join(clean(field(k)) for k in ('Str_num', 'Str_pfx', 'Str_name', 'Str_sfx') if field(k))
            display_number=re.match(r'\d+',address)
            description=clean(field('Property Description')) or f"{display_number[0] if display_number else '0'} ({legal})"
            rows.append({'row': rn, 'account': account, 'address': address, 'legal': legal,
                         'subdivision': clean(field('Legal_Dscr_2')), 'owner': clean(field('OWNER_MAILTO', 'OWNER_LIST')),
                         'mailing': clean(field('Mail_Addr_1')), 'area': number(field('Total_Land_Area', 'Lot Size')),
                         'land_use': clean(field('Land Use')).upper(), 'description': description,
                         'location': f'{s1.title}!row {rn}'})
        header = next((r for r in range(12, 21) if norm(s2.cell(r, 1).value) == 'PROPERTYDESCRIPTION'), None)
        if not header:
            raise ValueError('The Step 2 property table was not found.')
        old = renewal and 'most frequent' in clean(s2['D10'].value).lower() and not s2['D2'].value
        minimum_ref = 'E10' if old else 'D2' if renewal else 'L16' if family == 'mbl' else 'K16'
        minimum = number(value(s2.title, minimum_ref))
        if minimum is None and not renewal:
            m = re.match(r'([\d,]+(?:\.\d+)?)', clean(value(s2.title, 'E4')))
            minimum = number(m[1]) if m else None
        table = []
        for rn in range(header + 1, min(s2.max_row, 5000) + 1):
            vals = {c: value(s2.title, f'{c}{rn}') for c in 'ABCDEF'}
            link=re.fullmatch(r"='([^']+)'!\$?[A-Z]+\$?(\d+)",str(expressions[s2.title].get(f'A{rn}') or ''))
            source_row=int(link[2]) if link and link[1]==s1.title else None
            desc = clean(vals['A'])
            if not desc or desc == '0':
                # Standard direct links can have no formula cache in Step 1.
                formula = expressions[s2.title].get(f'A{rn}')
                m = re.fullmatch(r"='([^']+)'!\$?[A-Z]+\$?(\d+)", str(formula or ''))
                parent = next((p for p in rows if m and m[1] == s1.title and p['row'] == int(m[2])), None)
                if not parent:
                    continue
                desc = parent['description'] or f"{parent['address']} ({parent['legal']})"
            lu = clean(vals['B']).upper()
            area_col = 'E' if renewal and (not old or family=='mbl') else 'D'
            metric_col = 'D' if family == 'mbl' and renewal else 'E' if family == 'mbl' else area_col
            table.append({'row': rn, 'source_row':source_row, 'description': desc, 'land_use': lu,
                          'area': number(vals[area_col]), 'measure': number(vals[metric_col]),
                          'original': number(vals['C' if family == 'mbl' else 'D']) if renewal and not old else None,
                          'signed': clean(vals['C']).upper() if not renewal else None,
                          'flag': clean(vals['F']).upper() if renewal and not old else None,
                          'area_cell': f'{area_col}{rn}', 'measure_cell': f'{metric_col}{rn}',
                          'location': f'{s2.title}!row {rn}'})
        ci = {clean(a): b for a, b in book['Case Info'].iter_rows(max_col=2, values_only=True) if a} if 'Case Info' in book else {}
        errors, uncached = [], []
        for s in book:
            if not (s.title.startswith('Step') or s.title == 'Case Info'):
                continue
            for row in s.iter_rows(max_row=min(s.max_row, 5000), max_col=min(s.max_column, 65)):
                for c in row:
                    if c.data_type == 'e':
                        errors.append((f'{s.title}!{c.coordinate}', str(c.value)))
        for ref in (minimum_ref, 'D9' if renewal else 'H11', 'A13' if renewal else 'A15'):
            if value(s2.title, ref) is None:
                uncached.append(f'{s2.title}!{ref}')
        return {'wizard': wizard, 'family': family, 'renewal': renewal, 'old': old,
                'application': clean(s2['B2'].value), 'minimum': minimum, 'minimum_ref': f'{s2.title}!{minimum_ref}',
                'street': clean(s2['B4'].value), 'block': clean(s2['D4'].value),
                'cross': [clean(s2[c].value) for c in ('B5', 'D5')],
                'sides': [clean(s2[c].value).lower() for c in ('B6', 'D6') if clean(s2[c].value)],
                'case_info': ci, 'rows': rows, 'table': table, 'errors': errors, 'uncached': uncached,
                'step1': s1.title, 'step2': s2.title,
                'cached': {c: value(s2.title, c) for c in ('A9','A11','A13','A15','D9','H11','F8','B8','D8')},
                'threshold_formula': str(expressions[s2.title].get('L17' if family == 'mbl' else 'K17') or '')}
    finally:
        book.close(); formulas.close()


def ocr_page(page, cancel=None):
    """Use the wizard's existing Windows OCR; no cloud or language model."""
    from .paths import WIZARD_FILES_DIR
    script = WIZARD_FILES_DIR / 'assets/application_import/windows_ocr.ps1'
    shell = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'System32/WindowsPowerShell/v1.0/powershell.exe'
    if not shell.exists() or not script.exists():
        raise ValueError('Windows OCR is unavailable.')
    check_cancel(cancel)
    with tempfile.TemporaryDirectory(prefix='rca_peer_ocr_') as folder:
        png = Path(folder) / 'page.png'
        scale = min(3, 3500 / max(page.rect.width, page.rect.height))
        page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False).save(png)
        process = subprocess.Popen([str(shell), '-NoProfile', '-NonInteractive', '-File', str(script), '-ImagePath', str(png)],
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                   creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        import time
        start = time.monotonic()
        try:
            while process.poll() is None:
                check_cancel(cancel)
                if time.monotonic() - start > 55:
                    raise ValueError('Text recognition timed out on this page.')
                if cancel:
                    cancel.wait(.1)
                else:
                    time.sleep(.1)
        except BaseException:
            process.kill(); process.wait(); raise
        result = Path(str(png) + '.ocr.json')
        if process.returncode or not result.exists():
            raise ValueError('Windows OCR could not read this page.')
        data = json.loads(result.read_text(encoding='utf-8-sig'))
        return '\n'.join(' '.join(w['text'] for w in line['words']) for line in data['lines'])


def read_document(path, progress=None, cancel=None, use_ocr=True):
    """Return evidence blocks with real PDF pages or Word paragraph locations."""
    progress = progress or (lambda text: None)
    path = Path(path); suffix = path.suffix.lower()
    blocks = []
    if suffix == '.docx':
        with zipfile.ZipFile(io.BytesIO(read_shared_bytes(path))) as z:
            ns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            for part in ['word/document.xml'] + sorted(n for n in z.namelist() if re.fullmatch(r'word/(header|footer)\d+\.xml', n)):
                root = ET.fromstring(z.read(part))
                for i, p in enumerate(root.findall('.//w:p', ns), 1):
                    text = ''.join(t.text or '' for t in p.findall('.//w:t', ns))
                    if clean(text):
                        blocks.append({'location': f'{Path(part).stem}, paragraph {i}', 'text': text, 'method':'Word text'})
            images = [n for n in z.namelist() if n.startswith('word/media/') and n.lower().endswith(('.png','.jpg','.jpeg','.tif','.tiff','.bmp'))]
            for i, name in enumerate(images, 1):
                check_cancel(cancel)
                if not use_ocr:
                    blocks.append({'location':f'embedded image {i}', 'text':'', 'method':'unreadable', 'warning':'Image text was not checked (OCR is off).'}); continue
                progress(f'Reading {path.name}, embedded image {i} of {len(images)}...')
                try:
                    with fitz.open(stream=z.read(name)) as picture, fitz.open(stream=picture.convert_to_pdf(), filetype='pdf') as pdf:
                        if max(pdf[0].rect.width, pdf[0].rect.height) < 250:
                            continue
                        blocks.append({'location':f'embedded image {i}', 'text':ocr_page(pdf[0], cancel), 'method':'OCR'})
                except InterruptedError:
                    raise
                except Exception as error:
                    blocks.append({'location':f'embedded image {i}', 'text':'', 'method':'unreadable', 'warning':str(error)})
        return blocks
    if suffix == '.doc':
        from .office_convert import office_to_pdf
        with tempfile.TemporaryDirectory(prefix='rca_peer_word_') as temp:
            copy = Path(temp) / path.name; copy.write_bytes(read_shared_bytes(path))
            converted = office_to_pdf(copy, Path(temp) / 'pdf')
            return read_document(converted, progress, cancel, use_ocr)
    if suffix != '.pdf':
        raise ValueError('Select a PDF, DOCX, or DOC document.')
    with fitz.open(stream=read_shared_bytes(path), filetype='pdf') as doc:
        if doc.needs_pass:
            raise ValueError('This PDF is password protected.')
        if len(doc) > 500:
            raise ValueError('Select a PDF with 500 pages or fewer.')
        for i, page in enumerate(doc):
            check_cancel(cancel); progress(f'Reading {path.name}, page {i+1} of {len(doc)}...')
            native = page.get_text(sort=True)
            broken = (sum((ord(c)<32 and c not in '\n\r\t') or 127<=ord(c)<=159 or c=='\ufffd' for c in native) > max(3, len(native)*.025)
                      or sum(ord(c)>127 for c in native)>max(20,len(native)*.2))
            large_image = any(fitz.Rect(v['bbox']).get_area() > page.rect.get_area()*.25 for v in page.get_image_info())
            needs_ocr = broken or len(clean(native)) < 35 or large_image
            block = {'location':f'page {i+1}', 'page':i+1, 'text':native, 'method':'PDF text'}
            if needs_ocr:
                if use_ocr:
                    try:
                        block.update(text=ocr_page(page, cancel), method='OCR')
                    except InterruptedError:
                        raise
                    except Exception as error:
                        block.update(warning=str(error), method='partial' if not broken and clean(native) else 'unreadable')
                        if broken:
                            block['text']=''
                else:
                    block.update(warning='Scan or embedded image was not checked (OCR is off).', method='partial' if not broken and clean(native) else 'unreadable')
                    if broken:
                        block['text']=''
            blocks.append(block)
    return blocks


def hcad_record(text):
    """Only labelled HCAD fields. Monetary values and building SF are not land SF."""
    text = text.replace('\r',''); flat = clean(text)
    m = re.search(r'\bAccount\s*(?:Number|No\.?|#)?\s*:?\s*(\d[\d -]{10,20}\d)\b', flat, re.I)
    account = clean_account(m[1]) if m else ''
    if not account:
        return None
    def field(label, stop):
        match = re.search(r'(?:^|\n)\s*(?:' + label + r')\s*:?\s*\n?([\s\S]*?)(?=\n\s*(?:' + stop + r')\b|$)', text, re.I)
        return clean(match[1]) if match else ''
    area_match = re.search(r'\bLand\s*(?:Area|Size)?\s*:?\s*([\d,]+(?:\.\d+)?)\s*(?:SF|SQ\.?\s*F(?:T|EET)|Square\s+Feet)\b', flat, re.I)
    legal = field(r'Legal\s*(?:Description)?', r'Land|Building|Valuation|Location|Fiduciary|Market|Appraised|Tax')
    owner = field(r'(?:Owner\s*(?:Name)?|Name)', r'Mailing|Address|Owner|Legal|Land|Property|Valuation|Exemption')
    mailing = field(r'Mailing\s+Address', r'Valuation|Owner|Legal|Land|Property|Exemption|Jurisdictions?')
    # The current HCAD viewer can place labels in one column and values in
    # another. OCR emits its label column first. Require that exact label order
    # and a legal-description start before associating the first SF with Land.
    columns = re.search(r'Legal\s+Description\s+Land\s+(?:(?:Living\s+Area|Building\s+Area|Building\s+Class|Units|Net\s+Rentable\s+Area|Fiduciary|Status)\s+)*((?:LTS?|LOTS?|TRS?|TRACTS?|RES)\b.*?)\s+([\d, ]+(?:\.\d+)?)\s+SF\b',flat,re.I)
    if columns:
        legal=columns[1]
        if area_match is None:
            area_value=number(columns[2].replace(' ',''))
        else:area_value=number(area_match[1])
    else:
        area_value=number(area_match[1]) if area_match else None
        if norm(legal) in ('LANDLIVINGAREA','LAND','DESCRIPTIONLANDLIVINGAREA'):legal=''
    address = field(r'(?:Site|Property)\s+Address', r'Name|Owner|Mailing|Legal|Land|Valuation|Account')
    if not address:
        m = re.search(r'(?:^|\n)\s*(\d{1,6}\s+[A-Z][A-Z0-9 .-]{2,55})\s*\n\s*HOUSTON\s*[, ]\s*TX', text, re.I)
        address = clean(m[1]) if m else ''
    return {'account':account, 'area':area_value,
            'legal':legal, 'owner':owner, 'mailing':mailing, 'address':address}
