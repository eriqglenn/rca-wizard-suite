"""Prior MLS/MBL designations from public City records, without AI or guessed dates."""
from __future__ import annotations

import hashlib
import io
import json
import re
import time
import zipfile
from datetime import date, datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse, urlencode
from urllib.request import Request, urlopen

from lxml import etree as ET, html
from openpyxl import load_workbook
from shapely.ops import unary_union

from .boundary_map import PublicMapClient, parcel_polygon
from .council_district import district_polygon
from .workbooks import read_shared_bytes

BASE = 'https://mycity2.houstontx.gov/gisweb01/rest/services/HoustonMap/Planning_and_Development/MapServer/'
SERVICES = {'mls': BASE + '13', 'mbl': BASE + '12'}
PROGRAM_URL = 'https://www.houstontx.gov/planning/Min-Lot_Size-Min_Bldg_Line.html'
AGENDA_URL = 'https://www.houstontx.gov/citysec/agenda/agendaindex.html'
CASE_FIELDS = {
    'original_ordinance_no': 'Original_Ordinance_No',
    'original_passage_date': 'Original_Passage_Date',
    'original_effective_date': 'Original_Effective_Date',
    'original_expiration_date': 'Original_Expiration_Date',
}
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def ordinance(value):
    match = re.fullmatch(r'(?:ORD(?:INANCE)?\.?\s*(?:NO\.?\s*)?)?(\d{4})\s*[-–]\s*(\d{1,5})', str(value or '').strip(), re.I)
    return f'{match[1]}-{int(match[2]):04d}' if match else ''


def app_number(value):
    return re.sub(r'[^0-9A-Z]', '', re.sub(r'^(?:MLS|MBL)\s*', '', str(value or '').upper()))


def day(value):
    if isinstance(value, datetime): return value.date()
    if isinstance(value, date): return value
    if isinstance(value, (int, float)):
        # ArcGIS date-only fields are UTC midnight. Local conversion loses a day.
        if value > 10**10:
            try: return datetime.fromtimestamp(value / 1000, timezone.utc).date()
            except (ValueError, OverflowError, OSError): return None
        return None
    text = str(value or '').strip()
    for fmt in ('%m/%d/%Y', '%m.%d.%Y', '%Y-%m-%d', '%B %d, %Y', '%b %d, %Y', '%Y-%m-%d %H:%M:%S'):
        try: return datetime.strptime(text, fmt).date()
        except ValueError: pass
    return None


def date_text(value):
    result = day(value)
    return result.strftime('%m/%d/%Y') if result else ''


class CityRecords:
    """Download only public citywide data; match case locations locally."""
    def __init__(self, cache, progress=None, cancel=None, refresh=False):
        self.cache = Path(cache); self.cache.mkdir(parents=True, exist_ok=True)
        self.client = PublicMapClient(self.cache / 'queries', progress, cancel, use_cache=False)
        self.progress = self.client.progress
        self.refresh = refresh
        self.warnings = []
        self.sources = []

    def blob(self, url, limit=20_000_000):
        parsed = urlparse(url)
        if parsed.scheme != 'https' or parsed.hostname not in ('www.houstontx.gov', 'houstontx.gov'):
            raise ValueError('The designation source is not an official City URL.')
        target = self.cache / (hashlib.sha256(url.encode()).hexdigest() + '.dat')
        self.client.check_cancelled()
        if not self.refresh and target.exists() and time.time() - target.stat().st_mtime < 86400:
            return target.read_bytes()
        try:
            with urlopen(Request(url, headers={'User-Agent': 'Houston-RCA-Wizard/1.0'}), timeout=12) as response:
                raw = response.read(limit + 1)
            if len(raw) > limit: raise ValueError('City source exceeds the download size limit.')
            self.client.check_cancelled()
            temp = target.with_suffix('.tmp'); temp.write_bytes(raw); temp.replace(target)
            return raw
        except Exception:
            self.client.check_cancelled()
            if target.exists():
                checked = datetime.fromtimestamp(target.stat().st_mtime, timezone.utc).date()
                self.warnings.append(f'City lookup unavailable; using saved source from {checked}: {url}')
                return target.read_bytes()
            raise

    def designations(self, family):
        target = self.cache / (family + '_designations.json')
        saved = None
        if target.exists():
            try: saved = json.loads(target.read_text(encoding='utf8'))
            except (ValueError, OSError): pass
        if saved and not self.refresh and time.time() - saved['retrieved'] < 86400:
            self.sources.append(saved['source']); return saved['features']
        try:
            self.progress('Checking the City designation records...')
            rows = self.client.query(SERVICES[family], fields='*')
            source = {'url': SERVICES[family], 'checked_utc': datetime.now(timezone.utc).isoformat()}
            data = {'retrieved': time.time(), 'features': rows, 'source': source}
            temp = target.with_suffix('.tmp'); temp.write_text(json.dumps(data), encoding='utf8'); temp.replace(target)
            self.sources.append(source)
            return rows
        except Exception:
            self.client.check_cancelled()
            if not saved: raise
            self.sources.append(saved['source'])
            self.warnings.append('Live designation lookup unavailable; using City records saved ' + saved['source']['checked_utc'] + '.')
            return saved['features']

    def status_records(self, family):
        tree = html.fromstring(self.blob(PROGRAM_URL))
        links = [urljoin(PROGRAM_URL, a.get('href')) for a in tree.xpath('//a[@href]')
                 if '.xlsx' in a.get('href', '').lower() and ('update' in a.get('href', '').lower() or 'status' in a.text_content().lower())]
        if len(set(links)) != 1: raise ValueError('The City status workbook link could not be identified.')
        url = links[0]
        rows = status_rows(self.blob(url), family, url)
        self.sources.append({'url': url, 'checked_utc': datetime.now(timezone.utc).isoformat()})
        return rows

    def passage(self, number, hint):
        """Bounded fallback near the map's start date; require explicit adopted action."""
        hint = day(hint)
        if not hint: return None
        tree = html.fromstring(self.blob(AGENDA_URL))
        candidates = []
        for anchor in tree.xpath('//a[@href]'):
            if anchor.text_content().strip().lower() != 'minutes': continue
            url = urljoin(AGENDA_URL, anchor.get('href'))
            match = re.search(r'/(\d{8})\.pdf$', url, re.I)
            if not match: continue
            try: meeting = datetime.strptime(match[1], '%Y%m%d').date()
            except ValueError: continue
            if abs((meeting - hint).days) <= 8: candidates.append((abs((meeting - hint).days), url))
        for _, url in sorted(set(candidates))[:3]:
            import fitz
            raw = self.blob(url)
            with fitz.open(stream=raw, filetype='pdf') as pdf:
                text = '\n'.join(page.get_text() for page in pdf)
            value = adopted_date(text, number)
            if value:
                return {'value': value, 'source': url}
        return None


def status_rows(raw, family, url):
    """Read by headings; ignore wide formatted-but-empty columns in City workbooks."""
    book = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    result = []
    try:
        for sheet in book:
            title = sheet.title.lower()
            if family == 'mls' and not ('mls' in title or 'lot size' in title): continue
            if family == 'mbl' and not ('mbl' in title or 'building line' in title): continue
            headings = None
            for index, row in enumerate(sheet.iter_rows(max_col=24, max_row=min(sheet.max_row, 10000), values_only=True), 1):
                if index <= 12:
                    names = {re.sub(r'[^a-z]', '', str(v or '').lower()): i for i, v in enumerate(row)}
                    if 'ordinance' in names:
                        headings = names; continue
                if not headings: continue
                number = ordinance(row[headings['ordinance']])
                if not number: continue
                approval_col = headings.get('councilapprovaldate')
                result.append({'ordinance': number, 'passage': date_text(row[approval_col]) if approval_col is not None else '',
                               'application': str(row[headings.get('app', 0)] or ''),
                               'source': f'{url} | {sheet.title}, row {index}'})
    finally: book.close()
    return result


def adopted_date(text, number):
    """A reference to an old ordinance in meeting minutes is not its adoption."""
    year, sequence = ordinance(number).split('-')
    actions=list(re.finditer(r'\bORDINANCE\s+(?:NO\.?\s*)?' + year + r'\s*-\s*0*' + str(int(sequence)) + r'\s+(?:FINALLY\s+)?(?:ADOPTED|PASSED)\b', text, re.I))
    if not actions:return ''
    stamp=r'(?:TUESDAY|WEDNESDAY|THURSDAY|FRIDAY|MONDAY)\s*[-,]?\s*((?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\s+\d{1,2},?\s+\d{4})'
    # Two-day minutes start on Tuesday but ordinances normally appear after the
    # Wednesday reconvening. Use the session preceding the specific action.
    markers=[]
    for pattern in (r'^[ \t]*'+stamp, r'\b(?:convened|reconvened)\b[\s\S]{0,100}?'+stamp):
        for match in re.finditer(pattern,text,re.I|re.M):
            value=date_text(re.sub(r'(\d)\s+(\d{4})$',r'\1, \2',match[1]))
            if value:markers.append((match.start(),value))
    dates=set()
    for action in actions:
        previous=[m for m in markers if m[0]<action.start()]
        if not previous:return ''
        dates.add(max(previous)[1])
    return next(iter(dates)) if len(dates)==1 else ''


def match_designation(features, records, family, fields):
    parcels = [parcel_polygon(f) for f in features]
    if not parcels: return None, [], ['No selected parcels are available for the previous-designation lookup.']
    area = unary_union(parcels); candidates = []; warnings = []
    requested = ordinance(fields.get('original_ordinance_no'))
    current_app = app_number(fields.get('application_number'))
    cutoff = day(fields.get('date_received') or fields.get('application_filed_date')) or date.today()
    grouped = {}
    for raw in records:
        a = raw.get('attributes', {})
        number = ordinance(a.get('ORDINANCE') if family == 'mls' else a.get('ORDINANCE_'))
        try:
            polygon = district_polygon(raw)
            overlap = area.intersection(polygon).area
            if overlap <= max(1, area.area * .001): continue
        except Exception: continue
        status = str(a.get('Status') or '').strip().lower()
        kind = str(a.get('Type') or '').upper()
        if 'rescind' in status or 'resc' in kind.lower():
            warnings.append('The selected area intersects a rescinded designation. Verify its history before using renewal information.')
            continue
        if not number or status not in ('approved', 'expired'): continue
        effective = day(a.get('Effective_Date')) if family == 'mls' else None
        start = day(a.get('STARTDATE') if family == 'mls' else a.get('START_DATE'))
        record_app = app_number(a.get('App_Num'))
        if not requested:
            if (effective or start) and (effective or start) >= cutoff: continue
            if current_app.endswith('REN') and record_app == current_app and not (effective or start): continue
        grouped.setdefault((number, record_app), []).append({'attributes':a, 'polygon':polygon, 'start':date_text(start)})
    for (number, record_app), parts in grouped.items():
        polygon=unary_union([p['polygon'] for p in parts]);coverage=area.intersection(polygon).area/area.area
        each=min(p.intersection(polygon).area/p.area for p in parcels)
        candidates.append({'ordinance':number,'application':record_app,'coverage':coverage,
                           'complete':coverage>=.98 and each>=.95,'attributes':parts[0]['attributes'],
                           'polygon':polygon,'object_id':parts[0]['attributes'].get('OBJECTID'),
                           'start':parts[0]['start'],'parts':parts})
    public = [{k: v for k, v in c.items() if k not in ('polygon', 'attributes','parts')} for c in candidates]
    good = [c for c in candidates if c['complete']]
    if requested:
        good = [c for c in good if c['ordinance'] == requested]
        if not good:
            return None, public, warnings + ['The entered ordinance does not match a designation covering the selected parcels. Existing values were kept.']
    if not good:
        return None, public, warnings + ['No single previous designation covers the selected parcels. Review the boundary or enter the prior ordinance details in Case Info.']
    numbers = {c['ordinance'] for c in good}
    if len(numbers) > 1:
        # Select a later renewal only when the application lineage and footprint agree.
        bases = {re.sub(r'REN$', '', c['application']) for c in good}
        same_footprint = all(good[0]['polygon'].intersection(c['polygon']).area / good[0]['polygon'].union(c['polygon']).area >= .95 for c in good)
        if len(bases) != 1 or '' in bases or not same_footprint:
            return None, public, warnings + ['More than one previous ordinance covers this boundary: ' + ', '.join(sorted(numbers)) + '. Enter the intended original ordinance in Case Info, then refresh the map.']
        latest = max(numbers, key=lambda n: tuple(map(int, n.split('-'))))
        good = [c for c in good if c['ordinance'] == latest]
    chosen = good[0]
    chosen['records'] = [p for c in good for p in c['parts']]
    return chosen, public, warnings


def lookup(features, wizard, fields, cache, progress=None, cancel=None, refresh=False, repository=None):
    result = {'values': {}, 'warnings': [], 'sources': [], 'candidates': [], 'checked_utc': datetime.now(timezone.utc).isoformat()}
    if wizard not in ('mls_renewal', 'mbl_renewal'): return result
    repo = repository or CityRecords(cache, progress, cancel, refresh)
    try:
        family = wizard.split('_')[0]
        chosen, candidates, warnings = match_designation(features, repo.designations(family), family, fields)
        result.update(candidates=candidates, warnings=warnings)
        if chosen:
            result['match'] = {k: v for k, v in chosen.items() if k not in ('polygon', 'attributes', 'records','parts')}
            number = chosen['ordinance']; values = result['values']; values['original_ordinance_no'] = number
            source_url=SERVICES[family]+'/query?'+urlencode({'f':'pjson','where':f"OBJECTID={chosen['object_id']}",'outFields':'*','returnGeometry':'false'})
            result['sources'].append({'field': 'original_ordinance_no', 'source': source_url, 'value': number})
            def field(key, attribute):
                dates = {date_text(c['attributes'].get(attribute)) for c in chosen['records']}; dates.discard('')
                if len(dates) == 1:
                    value = dates.pop(); values[key] = value
                    result['sources'].append({'field': key, 'source': SERVICES[family] + ' | ' + attribute, 'value': value})
                elif len(dates) > 1: result['warnings'].append('The City records disagree on ' + key.replace('_', ' ') + '; no date was filled.')
            if family == 'mls': field('original_effective_date', 'Effective_Date')
            field('original_expiration_date', 'Expiration_Date')
            try:
                status = [r for r in repo.status_records(family) if r['ordinance'] == number]
                dates = {r['passage'] for r in status if r['passage']}
                if len(dates) == 1:
                    values['original_passage_date'] = dates.pop()
                    result['sources'].extend({'field': 'original_passage_date', 'source': r['source'], 'value': r['passage']} for r in status if r['passage'])
                elif len(dates) > 1: result['warnings'].append('City status records disagree on the Council passage date; no passage date was filled.')
                else:
                    passage = repo.passage(number, chosen['start'])
                    if passage:
                        values['original_passage_date'] = passage['value']
                        result['sources'].append({'field': 'original_passage_date', **passage})
            except Exception as exc:
                repo.client.check_cancelled()
                result['warnings'].append('Council passage date could not be checked: ' + str(exc))
            if values.get('original_effective_date') and values.get('original_expiration_date') and day(values['original_expiration_date']) <= day(values['original_effective_date']):
                values.pop('original_expiration_date')
                result['warnings'].append('The published expiration date precedes the effective date; expiration was left blank for review.')
            if values.get('original_passage_date') and values.get('original_effective_date') and day(values['original_passage_date'])>day(values['original_effective_date']):
                values.pop('original_effective_date')
                result['warnings'].append('The published effective date precedes Council passage; effective date was left blank for review.')
            for key in CASE_FIELDS:
                if key not in values: result['warnings'].append(CASE_FIELDS[key].replace('_', ' ') + ' was not verified. Enter it from the original ordinance if needed.')
    except Exception as exc:
        repo.client.check_cancelled()
        result['warnings'].append('Previous designation lookup unavailable. Existing information was kept: ' + str(exc))
    result['warnings'].extend(repo.warnings)
    result['records_sources'] = repo.sources
    result['warnings'] = list(dict.fromkeys(result['warnings']))
    return result


def merge_fields(existing, result):
    """Fill blanks only. Do not mix dates from a different ordinance with planner input."""
    fields = dict(existing); notes = []; filled = []
    old = str(fields.get('original_ordinance_no') or '').strip()
    found = result.get('values', {}).get('original_ordinance_no')
    if old and found and ordinance(old) != ordinance(found):
        return fields, [], ['The lookup found ordinance ' + found + ', but the entered ordinance was kept. No associated dates were changed.']
    for key, value in result.get('values', {}).items():
        if key not in CASE_FIELDS or not value: continue
        current = fields.get(key)
        if current is None or str(current).strip() == '': fields[key] = value; filled.append(key)
        else:
            equal = ordinance(current) == ordinance(value) if key.endswith('_no') else day(current) == day(value)
            if not equal: notes.append(f'{CASE_FIELDS[key].replace("_", " ")}: kept {current}; City lookup suggests {value}.')
    return fields, filled, notes


def workbook_fields(source):
    book = load_workbook(io.BytesIO(read_shared_bytes(source)), read_only=True, data_only=True)
    try:
        fields = {str(a or '').strip().replace(' ', '_').lower(): b for a, b in book['Case Info'].iter_rows(max_col=2, values_only=True)} if 'Case Info' in book.sheetnames else {}
        step = next(s for s in book if s.title.lower().startswith('step 2'))
        fields.update(application_number=step['B2'].value, date_received=step['B3'].value)
        return fields
    finally: book.close()


def workbook_bytes(raw, result):
    """Patch only Case Info. Preserve worksheets, formulas, artwork and typed dates."""
    from .application_import import sheet_paths, set_cell
    book = load_workbook(io.BytesIO(raw), read_only=True, data_only=False)
    try:
        sheet = book['Case Info']
        cells = {str(row[0].value or '').strip().replace(' ', '_').lower(): row[1] for row in sheet.iter_rows(max_col=2)}
        existing = {k: c.value for k, c in cells.items()}
        values, filled, notes = merge_fields(existing, result)
        changes = {CASE_FIELDS[k]: values[k] for k in filled}
        for key in filled:
            if key.endswith('_date'): changes[CASE_FIELDS[key]] = day(values[key])
        report = summary(result, notes=notes)
        changes.update(Renewal_Lookup_Notes=report[:20000])
        if result['sources']:
            changes['Renewal_Lookup_Source']='\n'.join(dict.fromkeys(s['source'] for s in result['sources']))[:20000]
        changes={k:v for k,v in changes.items() if existing.get(k.lower())!=v}
        if changes:changes['Renewal_Lookup_Checked']=result['checked_utc']
        for k in list(changes):
            cell = cells.get(k.lower())
            if cell is not None and cell.data_type == 'f':
                notes.append('Kept the formula in ' + k + '.'); changes.pop(k)
        last = sheet.max_row
        date_ref = next((c.coordinate for k, c in cells.items() if k.endswith('_date') and c.has_style), None)
    finally: book.close()
    if not changes:return raw,filled,notes
    with zipfile.ZipFile(io.BytesIO(raw)) as original:
        path = sheet_paths(original)['Case Info']; xml = ET.fromstring(original.read(path))
        date_cell=xml.find(f'.//s:c[@r="{date_ref}"]', NS) if date_ref else None
        date_style=date_cell.get('s') if date_cell is not None else None
        for label, value in changes.items():
            cell = cells.get(label.lower())
            if cell is None:
                last += 1; ref = f'B{last}'; set_cell(xml, f'A{last}', label)
                # Match existing Case Info inputs rather than restyling the sheet.
                template = xml.find('.//s:c[@r="B6"]', NS)
            else: ref = cell.coordinate; template = None
            set_cell(xml, ref, value)
            element = xml.find(f'.//s:c[@r="{ref}"]', NS)
            if isinstance(value, date) and date_style: element.set('s', date_style)
            elif template is not None and template.get('s'): element.set('s', template.get('s'))
        dimension = xml.find('s:dimension', NS)
        if dimension is not None:
            from openpyxl.utils.cell import range_boundaries, get_column_letter
            x0, y0, x1, y1 = range_boundaries(dimension.get('ref'))
            dimension.set('ref', f'{get_column_letter(x0)}{y0}:{get_column_letter(max(2,x1))}{max(last,y1)}')
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as output:
            for member in original.infolist(): output.writestr(member, ET.tostring(xml) if member.filename == path else original.read(member.filename))
    return buffer.getvalue(), filled, notes


def save_workbook(source, result, output=None, expected_hash=None):
    from .automatic_case import _commit, _digest
    source = Path(source); raw = read_shared_bytes(source)
    if expected_hash and _digest(raw) != expected_hash: raise ValueError('The workbook changed during designation lookup. Its renewal fields were not updated.')
    updated, filled, notes = workbook_bytes(raw, result)
    path, commit_notes = _commit(source, raw, updated, output)
    result['filled'] = filled; result['warnings'].extend(notes)
    return path, [summary(result, filled, notes)] + commit_notes


def summary(result, filled=None, notes=None):
    number = result.get('values', {}).get('original_ordinance_no')
    lines = ['Previous designation: ordinance ' + number + '.' if number else 'Previous designation was not automatically filled.']
    for key in ('original_passage_date', 'original_effective_date', 'original_expiration_date'):
        if result.get('values', {}).get(key): lines.append(CASE_FIELDS[key].replace('_', ' ') + ': ' + result['values'][key])
    if filled: lines.append('Filled blank Case Info fields: ' + ', '.join(CASE_FIELDS[k].replace('_', ' ') for k in filled) + '.')
    lines.extend(result.get('warnings', [])); lines.extend(notes or [])
    return '\n'.join(dict.fromkeys(lines))
