"""Reviewed application boundaries -> public HCAD facts -> City calculation copy.

The City's xlsx is patched at the cell XML level. All drawings (including WMF),
formulas, formatting, validations and other package parts remain intact.
"""
from __future__ import annotations

from datetime import datetime, date, timezone
import copy
import hashlib
import json
import math
from pathlib import Path
import re
import shutil
import tempfile
from types import SimpleNamespace
import zipfile

from lxml import etree as ET
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.datetime import to_excel
from shapely.geometry import LineString, Point, box
from shapely.ops import nearest_points, unary_union

from .boundary_map import (PARCEL_SERVICE, CITY_PARCEL_SERVICE, PublicMapClient, account_number, clean,
    road_name, road_key, street_identity, parcel_polygon, load_case_roads, MapError)
from .workbooks import calculation_template_path, is_bundled_input


LAND_USES = {'SFR':'SINGLE-Family Residential', 'MF':'Multi-Family Residential',
             'VAC':'Undeveloped', 'COM':'Commercial', 'EXC':'Excluded'}
WIZARD_LABELS = {'mls':'MLS', 'mbl':'MBL', 'mls_renewal':'MLS Renewal', 'mbl_renewal':'MBL Renewal'}
PROPERTY_INPUTS = {
    'mls': [('support','Signed support (reviewed)')],
    'mbl': [('support','Signed support (reviewed)'), ('building_line','Building line (feet)')],
    'mls_renewal': [('original_lot_area','Original lot area (sq ft; blank for newer lots)')],
    'mbl_renewal': [('original_building_line','Original building line (feet)'), ('building_line','Current building line (feet)')],
}


def wizard_label(wizard):
    if wizard not in WIZARD_LABELS: raise ValueError('Choose MLS, MBL, MLS Renewal or MBL Renewal.')
    return WIZARD_LABELS[wizard]


def case_number(value, wizard='mls'):
    label=wizard_label(wizard)
    match=re.fullmatch(r'(?:(MLS|MBL)\s*)?(\d{1,5})\s*(REN)?',clean(value),re.I)
    if not match or (match[1] and match[1].lower()!=wizard[:3]) or (match[3] and not wizard.endswith('_renewal')):
        raise ValueError(f'Enter the {label} application number using digits'+(' and an optional REN suffix.' if wizard.endswith('_renewal') else '.'))
    return match[2]+('REN' if wizard.endswith('_renewal') else '')


def optional_measurement(value, label, *, positive=False):
    if value is None or clean(value)=='': return None
    try: result=float(clean(value).replace(',',''))
    except ValueError: raise ValueError(f'Enter a number for {label}, or leave it blank for review.') from None
    if not math.isfinite(result) or result<0 or (positive and result==0):
        raise ValueError(f'Enter a {"positive" if positive else "nonnegative"} number for {label}.')
    return result


def calculation_review_notes(fields, selected, wizard):
    notes=[]
    if wizard.endswith('_renewal') and not clean(fields.get('renewal_minimum')):
        notes.append('Enter the existing minimum to renew in Step 2 D2 before creating the map.')
    if wizard.startswith('mbl') and any(not clean(p.get('building_line')) for p in selected):
        notes.append('Building-line measurements are blank where not supplied. Confirm them from the application or survey and complete Step 2; HCAD lot area cannot establish a setback.')
    if wizard.endswith('_renewal'):
        notes.append('Review original lot sizes/building lines and ordinance details against the prior designation. Current HCAD records do not establish historical values.')
    return notes
FIELDS = ('OBJECTID,HCAD_NUM,acct_num,tax_year,owner_name_1,owner_name_2,owner_name_3,'
          'mail_addr_1,mail_addr_2,mail_city,mail_state,mail_zip,site_str_pfx,site_str_num,'
          'site_str_num_sfx,site_str_name,site_str_sfx,site_str_sfx_dir,site_city,site_zip,'
          'state_class,land_value,bld_value,impr_value,productivity_value,total_appraised_val,'
          'total_market_val,new_owner_date,legal_dscr_1,legal_dscr_2,legal_dscr_3,legal_dscr_4,'
          'acreage_1,land_sqft,land_use,confidential_flag,activeAccount_flag,Stacked')


def numeric_range(text):
    match = re.fullmatch(r'\s*(\d{2,5})(?:\s*(?:-|–|and|to|through)\s*(\d{2,5}))?(?:\s+blocks?)?\s*', text, re.I)
    if not match: raise ValueError('Enter a block or address range, such as 2000, 2200-2300 or 4806-4834.')
    lo, hi = int(match[1]), int(match[2] or match[1])
    if hi < lo: raise ValueError('The ending block/address must be at least the starting number.')
    if lo % 100 == 0 and hi % 100 == 0: hi += 99
    if hi-lo > 3000: raise ValueError('Review the range: this importer handles a street segment up to 3,000 address numbers at a time.')
    return lo, hi


def validate_fields(fields, saving=False, wizard='mls'):
    wizard_label(wizard)
    required = ['street','block_range','cross_1','cross_2','sides']
    if saving: required += ['application_number']
    missing = [k.replace('_',' ') for k in required if not clean(fields.get(k))]
    if missing: raise ValueError('Complete these application fields: '+', '.join(missing)+'.')
    numeric_range(fields['block_range'])
    sides = re.findall(r'\b(north|south|east|west)\b', fields['sides'].lower())
    if not sides or len(sides)>2 or any(s not in ('north','south','east','west') for s in re.split(r'\s+and\s+', fields['sides'].lower())):
        raise ValueError('Choose north, south, east, west, north and south, or east and west.')
    if saving: case_number(fields['application_number'],wizard)
    if wizard.endswith('_renewal'): optional_measurement(fields.get('renewal_minimum'),'the minimum to renew',positive=True)
    for key in ['date_received','date_complete','application_date','original_passage_date','original_effective_date','original_expiration_date',
                'application_filed_date','application_accepted_date']:
        if fields.get(key): parse_date(fields[key])


def block_description(fields):
    low,high=numeric_range(fields['block_range'])
    first,last=(low//100)*100,(high//100)*100
    span=str(first) if first==last else f'{first}-{last}'
    return span+' block of '+fields['street']


def prepare_reviewed_save(fields, parcel_data, wizard):
    """Keep the planner's selected parcels; save incomplete facts as review notes."""
    wizard_label(wizard)
    fields, data = copy.deepcopy(fields), copy.deepcopy(parcel_data)
    notes = list(data.get('warnings', []))
    boundary_keys = ('street', 'block_range', 'sides', 'cross_1', 'cross_2')
    if data.get('fields') and any(fields.get(k) != data['fields'].get(k) for k in boundary_keys):
        notes.append('Boundary wording changed after lookup. Saved the parcels accepted by the planner without repeating or changing the selection.')
    for key in boundary_keys + ('application_number',):
        fields.setdefault(key, '')
        if not clean(fields[key]):
            notes.append(f'Complete {key.replace("_", " ")} in the calculation sheet.')
    checks = [('application_number', lambda v: case_number(v, wizard)),
              ('block_range', numeric_range)]
    if wizard.endswith('_renewal'):
        checks.append(('renewal_minimum', lambda v: optional_measurement(v, 'minimum to renew', positive=True)))
    checks += [(k, parse_date) for k in ('date_received', 'date_complete', 'application_date',
               'original_passage_date', 'original_effective_date', 'original_expiration_date', 'application_filed_date', 'application_accepted_date')]
    for key, check in checks:
        if not clean(fields.get(key)):
            continue
        try:
            check(fields[key])
        except ValueError:
            notes.append(f'{key.replace("_", " ").title()} {fields[key]!r} needs review; its calculation cell was left blank.')
            fields[key] = ''
    for p in data['parcels']:
        if not p['include']:
            continue
        label = p.get('address') or p['account']
        if p.get('land_use') not in LAND_USES:
            notes.append(f'{label}: land use is unconfirmed; complete the classification in Step 1.')
            p['land_use'] = ''
        try:
            area = optional_measurement(p.get('land_area'), 'lot area', positive=True)
        except ValueError:
            area = None
        if area is None:
            notes.append(f'{label}: lot area is unconfirmed; its value was left blank in Step 1.')
        p['land_area'] = area
        if p.get('support', '') not in ('', 'Y', 'N'):
            notes.append(f'{label}: signed support is unconfirmed; its value was left blank.')
            p['support'] = ''
        for key, title in PROPERTY_INPUTS[wizard]:
            if key == 'support':
                continue
            try:
                p[key] = optional_measurement(p.get(key), title)
            except ValueError:
                notes.append(f'{label}: {title} needs review; its value was left blank.')
                p[key] = None
    data['warnings'] = list(dict.fromkeys(notes))
    return fields, data


def parse_date(value):
    if isinstance(value, (date,datetime)): return value
    for fmt in ('%m/%d/%Y','%m/%d/%y','%Y-%m-%d'):
        try: return datetime.strptime(clean(value),fmt)
        except ValueError: pass
    raise ValueError(f'Check the date {value!r}. Use MM/DD/YYYY.')


def site_address(a):
    return ' '.join(clean(a.get(k)) for k in ('site_str_num','site_str_num_sfx','site_str_pfx','site_str_name','site_str_sfx','site_str_sfx_dir') if clean(a.get(k)) not in ('','0'))


def suggested_land_use(a):
    # HCAD land codes are valuation codes, not City's case classifications.
    # Only clear state categories receive a suggestion; all remain editable.
    return {'A1':'SFR','B1':'MF','C1':'VAC','F1':'COM'}.get(clean(a.get('state_class')).upper(),'')


def parcel_row(feature, hints=()):
    a = feature['attributes']
    account = account_number(a.get('HCAD_NUM') or a.get('acct_num'))
    if not account: raise MapError('HCAD account is missing from this public record.')
    area = a.get('land_sqft')
    notes = []
    if not isinstance(area,(float,int)) or not math.isfinite(area) or area<=0:
        area = None; notes.append('HCAD land area is missing; confirm the record.')
    matching = [h for h in hints if h['account']==account]
    areas = sorted(set(h['area'] for h in matching if h.get('area') is not None))
    if area is not None and any(abs(v-area)>.001 for v in areas):
        notes.append('Application table: '+', '.join(f'{v:,.2f}'.rstrip('0').rstrip('.') for v in areas)+' sq ft; HCAD: '+f'{area:,.2f}'.rstrip('0').rstrip('.')+' sq ft.')
    if matching and any(h.get('number') and h['number']!=clean(a.get('site_str_num')) for h in matching):
        notes.append('Application account/address differs from HCAD. Check the application table.')
    lu = suggested_land_use(a)
    if not lu: notes.append('Choose the City land-use classification after reviewing the record.')
    if a.get('Stacked'): notes.append('Stacked parcel: review for duplicate land area or condominium accounts.')
    if clean(a.get('activeAccount_flag')).upper() not in ('','Y','1','TRUE'): notes.append('HCAD account is not marked active; confirm the record.')
    if not clean(a.get('owner_name_1')) or not clean(a.get('mail_addr_1')): notes.append('Owner or mailing information is missing in public data.')
    return {'account':account,'address':site_address(a),'land_area':area,'land_use':lu,
            'area_source':{'service':PARCEL_SERVICE,'field':'land_sqft','tax_year':clean(a.get('tax_year'))},
            'owner':clean(a.get('owner_name_1')),'mail_addr_1':clean(a.get('mail_addr_1')),
            'mail_addr_2':clean(a.get('mail_addr_2')),'mail_city':clean(a.get('mail_city')),
            'mail_state':clean(a.get('mail_state')),'mail_zip':clean(a.get('mail_zip')),
            'tax_year':clean(a.get('tax_year')), 'legal_description':clean(a.get('legal_dscr_1')),
            'subdivision':clean(a.get('legal_dscr_2')), 'include':False, 'support':'',
            'notes':notes,'review_note':'','record_pdfs':[], 'feature':feature}


def enrich_city_records(rows,client,hints=()):
    """Use the explicit Total_Land_Area field, keeping its separate source year.

    County land_sqft can reflect rounded acreage; never use polygon geometry
    area as a substitute for the appraisal record's land area.
    """
    client.progress('Checking reported land areas and property details in the City HCAD layer...')
    found={}
    # URL-encoded IN clauses with 100 accounts exceed this service's query
    # string limit and return HTTP 404. Keep each request below that limit.
    for i in range(0,len(rows),50):
        accounts=','.join("'"+r['account']+"'" for r in rows[i:i+50])
        for feature in client.query(CITY_PARCEL_SERVICE,f'HCAD_NUM IN ({accounts})',fields='*'):
            a=feature['attributes'];found.setdefault(clean(a.get('HCAD_NUM')),[]).append(a)
    for row in rows:
        matches=found.get(row['account'],[])
        if len(matches)!=1:continue
        a=matches[0];area=a.get('Total_Land_Area')
        row['secondary_attributes']=a
        if isinstance(area,(int,float)) and math.isfinite(area) and area>0:
            old=row['land_area'];year=clean(a.get('Tax_Year'))
            row['land_area']=area
            row['area_source']={'service':CITY_PARCEL_SERVICE,'field':'Total_Land_Area','tax_year':year}
            row['notes']=[s for s in row['notes'] if not s.startswith(('HCAD land area is missing','Application table:'))]
            if old is None or abs(old-area)>.001:
                row['notes'].append(f'Using reported land area {area:,.0f} sq ft (City HCAD data, {year}); county {row["tax_year"]} area is {old if old is not None else "missing"}. Confirm against the property record.')
            elif year!=row['tax_year']:
                row['notes'].append(f'Reported land-area source: City HCAD data, {year}. Owner source year: {row["tax_year"]}.')
            other=sorted({h['area'] for h in hints if h['account']==row['account'] and h.get('area') is not None and abs(h['area']-area)>.001})
            if other:row['notes'].append('Application table differs from reported land area: '+', '.join(str(v) for v in other)+' sq ft. Review both sources.')
        if not row['land_use']:
            key=clean(a.get('group_dscr')).lower()
            row['land_use']=next((code for code,desc in LAND_USES.items() if desc.lower()==key),'')
            if row['land_use']:row['notes'].append('Land-use suggestion comes from the City HCAD layer; confirm the current use.')


def _import_road_shape(roads, name):
    parts = [LineString(path) for r in roads
             if road_key(r['attributes'].get('FULL_NAME')) == road_key(name)
             for path in r.get('geometry', {}).get('paths', []) if len(path) > 1]
    return unary_union(parts) if parts else None


def _import_endpoints(fields, roads):
    main = _import_road_shape(roads, fields['street'])
    endpoints = []
    if main is not None:
        for key in ('cross_1', 'cross_2'):
            cross = _import_road_shape(roads, fields[key])
            if cross is None or main.distance(cross) > 100:
                return []
            endpoints.append(nearest_points(main, cross)[0])
    return endpoints if len(endpoints) == 2 and endpoints[0].distance(endpoints[1]) >= 30 else []


def _localize_import_parcels(fields, features, client, warnings):
    """Resolve countywide names against both requested cross streets, not mailing ZIPs."""
    polygons = [parcel_polygon(f) for f in features]
    joined = unary_union([p.buffer(600) for p in polygons])
    regions = list(joined.geoms) if joined.geom_type == 'MultiPolygon' else [joined]
    case = SimpleNamespace(street=fields['street'], cross_streets=[fields['cross_1'], fields['cross_2']])
    matches = []
    for region in regions:
        client.check_cancelled()
        group = [f for f, p in zip(features, polygons) if region.intersects(p)]
        bounds = unary_union([parcel_polygon(f) for f in group]).bounds
        if max(bounds[2]-bounds[0], bounds[3]-bounds[1]) > 12000:
            continue
        extent = (bounds[0]-160, bounds[1]-160, bounds[2]+160, bounds[3]+160)
        group_warnings = []
        client.progress('Checking the requested cross streets to distinguish same-name streets in different areas...')
        roads = load_case_roads(case, client, extent, group_warnings)
        if _import_endpoints(fields, roads):
            matches.append((group, roads, extent, group_warnings))
    if len(matches) != 1:
        raise ValueError('HCAD returned same-name addresses in separate areas. The requested cross streets '
                         'did not identify one unique area. Check both cross streets and retry; no area was guessed.')
    group, roads, extent, group_warnings = matches[0]
    warnings.extend(group_warnings)
    selected_ids = {f['attributes']['OBJECTID'] for f in group}
    outside = [f for f in features if f['attributes']['OBJECTID'] not in selected_ids]
    cities = sorted({clean(f['attributes'].get('site_city')) for f in outside if clean(f['attributes'].get('site_city'))})
    location = ' ('+', '.join(cities)+')' if cities else ''
    warnings.append(f'Used {fields["cross_1"]} and {fields["cross_2"]} to locate the requested street segment. '
                    f'{len(outside)} same-name address record(s) in other areas{location} were not included. '
                    'Review the local parcels before saving.')
    return group, roads, extent


def find_parcels(fields, cache_dir, hints=(), progress=None, cancel=None, use_cache=False):
    validate_fields(fields)
    client = PublicMapClient(cache_dir, progress, cancel, use_cache=use_cache)
    lo, hi = numeric_range(fields['block_range'])
    # Prefix/direction are checked after lookup, so "W Main" and "Main" never merge.
    tokens = street_identity(fields['street']).split()
    if tokens and tokens[0] in ('NORTH','SOUTH','EAST','WEST'): tokens = tokens[1:]
    base = ' '.join(tokens).replace("'","''")
    where = f"UPPER(site_str_name) = '{base}' AND site_str_num >= {lo} AND site_str_num <= {hi}"
    client.progress('Finding HCAD parcels for the application street and range...')
    initial = client.query(PARCEL_SERVICE,where,fields=FIELDS)
    initial = [f for f in initial if road_key(street_identity(' '.join(clean(f['attributes'].get(k)) for k in ('site_str_pfx','site_str_name','site_str_sfx','site_str_sfx_dir')))) == road_key(street_identity(fields['street']))]
    if not initial: raise ValueError('No HCAD parcels matched that street and range. Check the spelling and address range, then retry.')
    polys = [parcel_polygon(f) for f in initial]
    bounds = unary_union(polys).bounds
    warnings = []
    if max(bounds[2]-bounds[0],bounds[3]-bounds[1]) > 12000:
        initial, roads, extent = _localize_import_parcels(fields, initial, client, warnings)
    else:
        extent = (bounds[0]-160,bounds[1]-160,bounds[2]+160,bounds[3]+160)
        roads = load_case_roads(SimpleNamespace(street=fields['street'],cross_streets=[fields['cross_1'],fields['cross_2']]),client,extent,warnings)
    main = _import_road_shape(roads, fields['street'])
    endpoints = _import_endpoints(fields, roads)
    if not endpoints:
        warnings.append('Cross-street limits could not be confirmed automatically. Review the selected parcels on the map.')
    # Adjacent properties can have corner-street addresses or no assigned address.
    nearby = client.query(PARCEL_SERVICE,extent=extent,fields=FIELDS)
    all_features = {f['attributes']['OBJECTID']:f for f in [*initial,*nearby]}
    initial_ids = {f['attributes']['OBJECTID'] for f in initial}
    result, seen, unresolved = [], set(), []
    sides = fields['sides'].lower().split(' and ')
    for f in all_features.values():
        try:
            row = parcel_row(f,hints)
        except MapError as error:
            address=site_address(f['attributes']) or 'unaddressed parcel'
            unresolved.append({'feature':f,'address':address,'reason':str(error),
                'in_requested_street_range':f['attributes']['OBJECTID'] in initial_ids})
            warnings.append(f'HCAD record {f["attributes"]["OBJECTID"]} ({address}) needs review: {error} '
                'It is not in the calculation selection. Confirm its account and use Add HCAD account if it belongs in the boundary. Other records were loaded.')
            continue
        if row['account'] in seen:
            warnings.append(f"Multiple parcel features for {row['account']}; review that account before import.")
            continue
        seen.add(row['account'])
        center = parcel_polygon(f).representative_point()
        on_side = True
        if len(sides)==1 and main is not None:
            near = nearest_points(main,center)[0]
            delta = {'north':center.y-near.y,'south':near.y-center.y,'east':center.x-near.x,'west':near.x-center.x}[sides[0]]
            on_side = delta > 0
        elif main is None and len(sides)==1:
            on_side=False
        between=True
        if endpoints:
            p,q=endpoints;dx=q.x-p.x;dy=q.y-p.y
            t=((center.x-p.x)*dx+(center.y-p.y)*dy)/(dx*dx+dy*dy)
            between=0<=t<=1
        row['include'] = f['attributes']['OBJECTID'] in initial_ids and on_side and between
        if f['attributes']['OBJECTID'] not in initial_ids:
            row['notes'].append('Nearby parcel: include only if it belongs in the application boundary.')
        result.append(row)
    if main is None: warnings.append('Street geometry is unavailable. Select the appropriate side manually in the parcel list.')
    hint_accounts = {h['account'] for h in hints}
    outside = hint_accounts - {r['account'] for r in result if r['include']}
    if outside: warnings.append(f'The scanned property tables contain {len(outside)} account(s) outside the suggested selection or not matched by OCR. Review their pages; tables and signatures do not automatically expand the boundary.')
    try:enrich_city_records(result,client,hints)
    except ValueError as e:
        client.check_cancelled()
        warnings.append('The reported land-area lookup is unavailable. Review county values and any missing areas against HCAD PDFs. '+str(e))
    result.sort(key=lambda r:(road_key(r['feature']['attributes'].get('site_str_name')),int(r['feature']['attributes'].get('site_str_num') or 0),r['account']))
    return {'parcels':result,'roads':roads,'warnings':list(dict.fromkeys(warnings)),'unresolved_records':unresolved,
            'sources':client.sources,'fields':dict(fields),'extent':extent}


def lookup_account(account, cache_dir, hints=(), progress=None, cancel=None):
    account=account_number(account)
    client=PublicMapClient(cache_dir,progress,cancel,use_cache=False)
    features=client.query(PARCEL_SERVICE,f"HCAD_NUM = '{account}'",fields=FIELDS)
    if len(features)!=1: raise ValueError('That account did not return exactly one parcel. Check the account in HCAD.')
    row=parcel_row(features[0],hints);row['include']=True
    try:enrich_city_records([row],client,hints)
    except ValueError as e:row['notes'].append('Reported land-area lookup unavailable: '+str(e))
    return row,client.sources


def record_pdf_facts(path, progress=None, cancel=None):
    """Read labelled fields only from a downloaded HCAD PDF, for review."""
    import fitz
    with fitz.open(path) as doc:
        text='\n'.join(p.get_text() for p in doc)
    def facts(value):
        # A filename or a random number elsewhere in the packet cannot prove
        # this is the requested property record.
        labelled=re.search(r'\bAccount(?:\s+(?:Number|No\.?))?\s*[:#]?\s*(\d{13})\b',value,re.I)
        area=re.search(r'\bLand(?:\s+Area|\s+Size)?\s*:?\s*([\d,]+(?:\.\d+)?)\s*(?:SF|sq\.?\s*ft|square\s*feet)\b',value,re.I)
        return (labelled[1] if labelled else '',float(area[1].replace(',','')) if area else None)
    account,area=facts(text);method='PDF text'
    if not account and not re.search(r'No Results Found|Verify you are human|Checking your browser',text,re.I):
        from .application_ocr import read_application, read_small_account_text
        read=read_application(path,progress,cancel,force_ocr=True,page_limit=2)
        text='\n'.join(p['text'] for p in read['pages'])
        account,area=facts(text);method='Windows OCR'
        if not account:
            text+='\n'+read_small_account_text(path,read['pages'],cancel)
            account,area=facts(text)
    return {'account':account,'land_area':area,'text':text,'method':method,'path':str(Path(path).resolve())}


NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
N={'s':NS}
def tag(name): return '{'+NS+'}'+name


def set_cell(root, ref, value):
    data=root.find('s:sheetData',N)
    rn,cn=coordinate_to_tuple(ref)
    row=next((r for r in data if int(r.get('r'))==rn),None)
    if row is None:
        row=ET.Element(tag('row'),r=str(rn));data.append(row)
        data[:]=sorted(data,key=lambda r:int(r.get('r')))
    cell=next((c for c in row if c.get('r')==ref),None)
    if cell is None:
        cell=ET.Element(tag('c'),r=ref);row.append(cell)
        row[:]=sorted(row,key=lambda c:coordinate_to_tuple(c.get('r'))[1])
    for child in list(cell): cell.remove(child)
    cell.attrib.pop('t',None)
    if value is None or value=='': return
    if isinstance(value,(date,datetime)):
        value=to_excel(value)
    if isinstance(value,(int,float)):
        if not math.isfinite(value): raise ValueError('A workbook value is not a finite number.')
        ET.SubElement(cell,tag('v')).text=str(value)
    else:
        cell.set('t','inlineStr')
        text=ET.SubElement(ET.SubElement(cell,tag('is')),tag('t'))
        text.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
        text.text=str(value)


def sheet_paths(z):
    wb=ET.fromstring(z.read('xl/workbook.xml'))
    rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    mapping={r.get('Id'):r.get('Target') for r in rels}
    result={}
    for s in wb.find('s:sheets',N):
        target=mapping[s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')]
        result[s.get('name')]=target.lstrip('/') if target.startswith('/') else 'xl/'+target
    return result


def workbook_values(row):
    a=row['feature']['attributes']
    values={'TAX_ID':row['account'],'ALT_TAX_ID':row['account'],'HCAD_NUM_1':row['account'],
        'OWNER_MAILTO':row['owner'],'OWNER_LIST':row['owner'], 'Site_addr_1':row['address'],
        'Site_addr_2':a.get('site_city'),'Site_addr_3':a.get('site_zip'),
        'Total_Land_Area':row['land_area'],'group_dscr':LAND_USES.get(row['land_use'], ''),
        'Mail_Addr_1':row['mail_addr_1'],'Mail_Addr_2':row['mail_addr_2'],'Mail_City':row['mail_city'],
        'Mail_State':row['mail_state'],'Mail_Zip':row['mail_zip'],'COUNTY':'HARRIS',
        'Legal_Dscr_1':row['legal_description'],'Legal_Dscr_2':row['subdivision']}
    mapping={'OBJECTID':'OBJECTID','Tax_Year':'tax_year','Str_num':'site_str_num','str_num_sfx':'site_str_num_sfx',
      'Str_pfx':'site_str_pfx','Str_name':'site_str_name','Str_sfx':'site_str_sfx','Str_sfx_dir':'site_str_sfx_dir',
      'State_Class':'state_class','Acreage':'acreage_1','Land_Value':'land_value','Improvement_Value':'bld_value',
      'Extra_features_Value':'impr_value','Ag_Value':'productivity_value','Total_Appraised_Value':'total_appraised_val',
      'Total_Market_Value':'total_market_val','Legal_Dscr_3':'legal_dscr_3','Legal_Dscr_4':'legal_dscr_4',
      'landuse_cd':'land_use','Stacked':'Stacked'}
    values.update({k:a.get(v) for k,v in mapping.items()})
    secondary=row.get('secondary_attributes',{})
    for key in ('Yr_Impr','Total_Building_Area','group_cd','landuse_dscr'):
        if secondary.get(key) is not None:values[key]=secondary[key]
    if row.get('lot_id'): values['Acreage']=row['land_area']/43560 if row.get('land_area') else None
    if a.get('new_owner_date'):
        values['New_Owner_Date']=datetime.fromtimestamp(a['new_owner_date']/1000,timezone.utc).replace(tzinfo=None)
    return values


def save_calculation(destination, fields, parcel_data, application, wizard='mls', *, reviewed=False):
    submitted_fields, submitted_parcels = fields, parcel_data['parcels']
    if reviewed:
        fields, parcel_data = prepare_reviewed_save(fields, parcel_data, wizard)
    else:
        validate_fields(fields,saving=True,wizard=wizard)
    destination=Path(destination).resolve()
    if destination.suffix.lower()!='.xlsx' or is_bundled_input(destination):
        raise ValueError('Choose a new .xlsx file in your case folder, outside the bundled templates.')
    if destination.exists(): raise ValueError('That file already exists. Choose a new name to preserve the existing workbook.')
    from .lot_divisions import expand_parcels
    parcel_data=copy.deepcopy(parcel_data)
    parcel_data['parcels']=expand_parcels(parcel_data['parcels'])
    selected=[p for p in parcel_data['parcels'] if p['include']]
    for p in selected:
        if p.get('lot_id') and p.get('land_area') is None:parcel_data.setdefault('warnings',[]).append(p['address']+': individual lot area is unknown for '+p['legal_description']+'. Complete Step 1 before relying on the calculations.')
    if not selected: raise ValueError('Select at least one parcel.')
    for p in selected:
        account_number(p['account'])
        if not reviewed and p['land_use'] not in LAND_USES: raise ValueError(f"Choose the land-use classification for {p['address'] or p['account']}.")
        if not reviewed and (not isinstance(p['land_area'],(int,float)) or not math.isfinite(p['land_area']) or p['land_area']<=0):
            raise ValueError(f"Confirm a positive lot area for {p['address'] or p['account']}.")
        if p.get('support','') not in ('','Y','N'): raise ValueError('Support must be blank, Y or N.')
        for key,label in PROPERTY_INPUTS[wizard]:
            if key!='support': optional_measurement(p.get(key),label)
    template=calculation_template_path(wizard)
    source_dir=destination.with_name(destination.stem+'_sources')
    if source_dir.exists(): raise ValueError('The source folder for this name already exists. Choose another workbook name.')
    with zipfile.ZipFile(template) as src:
        paths=sheet_paths(src)
        s1name=next(k for k in paths if k.startswith('Step 1'))
        s2name=next(k for k in paths if k.startswith('Step 2'))
        roots={k:ET.fromstring(src.read(paths[k])) for k in (s1name,s2name,'Case Info')}
        s1,s2,ci=roots[s1name],roots[s2name],roots['Case Info']
        first,last=(17,125) if wizard=='mls' else ((17,99) if wizard=='mbl' else (15,97))
        capacity=last-first+1
        if len(selected)>capacity:
            raise ValueError(f'This City calculation template supports {capacity} properties in all its calculation ranges. This selection contains {len(selected)}; use an expanded calculation template for this case.')
        wb=load_workbook(template,read_only=True,data_only=False)
        try:
            header={clean(c.value).rstrip(' *'):c.column_letter for c in next(wb[s1name].iter_rows(min_row=1,max_row=1)) if c.value}
            case_cells={}
            static={'Planner_Position','Director_Name','PC_Signer','Ordinance_Year','Renewal_Term_Years'}
            for row in wb['Case Info'].iter_rows(min_row=6,max_col=2):
                label,c=row
                if c.data_type=='f': continue
                case_cells[clean(label.value)]=c.coordinate
                if clean(label.value) not in static: set_cell(ci,c.coordinate,None)
        finally: wb.close()
        for c in s1.findall('.//s:sheetData/s:row/s:c',N):
            if coordinate_to_tuple(c.get('r'))[0]>1 and c.find('s:f',N) is None:
                for child in list(c):c.remove(child)
                c.attrib.pop('t',None)
        for i,p in enumerate(selected,2):
            for key,value in workbook_values(p).items():
                if key in header: set_cell(s1,header[key]+str(i),value)
        inputs={'B2':case_number(fields['application_number'],wizard) if fields.get('application_number') else None,'B3':parse_date(fields['date_received']) if fields.get('date_received') else None,
          'D3':parse_date(fields['date_complete']) if fields.get('date_complete') else None,
          'B4':fields['street'],'D4':block_description(fields) if fields.get('block_range') else None,
          'B5':fields['cross_1'],'D5':fields['cross_2']}
        sides=fields['sides'].split(' and ');inputs.update(B6=sides[0].title(),D6=sides[1].title() if len(sides)>1 else '')
        if wizard.endswith('_renewal'): inputs['D2']=optional_measurement(fields.get('renewal_minimum'),'the minimum to renew',positive=True)
        for ref,value in inputs.items(): set_cell(s2,ref,value)
        # The template has 109 property formulas, but its two older SUM ranges
        # stopped at 98/99. Include every supported row without changing the
        # City's calculation method. No minimum is computed by the importer.
        if not wizard.endswith('_renewal'):
            for ref,column in [('B8','D'),('D8','Y' if wizard=='mls' else 'Z')]:
                cell=s2.find(f'.//s:c[@r="{ref}"]',N)
                cell.find('s:f',N).text=f"SUM('{s2name}'!${column}{first}:${column}{last})"
        property_columns={'mls':{'C':'support'}, 'mbl':{'C':'support','E':'building_line'},
            'mls_renewal':{'C':None,'D':'original_lot_area'},
            'mbl_renewal':{'C':'original_building_line','D':'building_line'}}[wizard]
        for rn in range(first,last+1):
            p=selected[rn-first] if rn-first<len(selected) else {}
            for column,key in property_columns.items():
                value=p.get(key) if key else None
                if key and key!='support': value=optional_measurement(value,key.replace('_',' '))
                set_cell(s2,column+str(rn),value)
        # Match Case Info by the City's field names; row positions differ in each wizard.
        case={'Planner_Name':'planner_name','Council_District':'council_district','Council_Member':'council_member','Applicant_Name':'applicant_name',
          'Applicant_Address_1':'applicant_address','Applicant_Address_2':'applicant_address_2',
          'Applicant_City':'applicant_city','Applicant_State':'applicant_state','Applicant_Zip':'applicant_zip',
          'Applicant_Property_Address':'applicant_property_address','Original_Ordinance_No':'original_ordinance_no',
          'Original_Passage_Date':'original_passage_date',
          'Original_Effective_Date':'original_effective_date','Original_Expiration_Date':'original_expiration_date',
          'Application_Filed_Date':'application_filed_date','Application_Accepted_Date':'application_accepted_date',
          'Application_Complete_Date':'date_complete'}
        if wizard.endswith('_renewal') and 'Original_Passage_Date' not in case_cells:
            next_row=max(int(r.get('r')) for r in ci.findall('s:sheetData/s:row',N))+1
            set_cell(ci,f'A{next_row}','Original_Passage_Date')
            passage_ref=f'B{next_row}';set_cell(ci,passage_ref,None);case_cells['Original_Passage_Date']=passage_ref
            style_cell=ci.find(f'.//s:c[@r="{case_cells.get("Original_Effective_Date", "B6")}"]',N)
            if style_cell is not None and style_cell.get('s'):ci.find(f'.//s:c[@r="{passage_ref}"]',N).set('s',style_cell.get('s'))
            dim=ci.find('s:dimension',N)
            if dim is not None:
                from openpyxl.utils.cell import range_boundaries,get_column_letter
                x0,y0,x1,y1=range_boundaries(dim.get('ref'));dim.set('ref',f'{get_column_letter(x0)}{y0}:{get_column_letter(max(x1,2))}{max(y1,next_row)}')
        for label,key in case.items():
            if label not in case_cells: continue
            value=fields.get(key,'')
            if value and label.endswith('_Date'): value=parse_date(value)
            set_cell(ci,case_cells[label],value)
        destination.parent.mkdir(parents=True,exist_ok=True)
        with tempfile.TemporaryDirectory(prefix='rca_import_save_',dir=destination.parent) as temp:
            tmp=Path(temp);out=tmp/destination.name;sources=tmp/'sources';sources.mkdir()
            with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as dst:
                for member in src.infolist():
                    payload=src.read(member.filename)
                    if member.filename in paths.values():
                        name=next(k for k,v in paths.items() if v==member.filename)
                        root=roots.get(name)
                        if root is None: root=ET.fromstring(payload)
                        for c in root.findall('.//s:c[s:f]',N):
                            v=c.find('s:v',N)
                            if v is not None: c.remove(v)
                        payload=ET.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
                    elif member.filename=='xl/workbook.xml':
                        root=ET.fromstring(payload);calc=root.find('s:calcPr',N)
                        if calc is None: calc=ET.SubElement(root,tag('calcPr'))
                        calc.set('calcMode','auto');calc.set('fullCalcOnLoad','1');calc.set('forceFullCalc','1')
                        payload=ET.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
                    dst.writestr(member,payload)
            if any(p.get('lot_id') for p in selected):
                from .lot_divisions import put_metadata,lot_key
                with zipfile.ZipFile(out) as saved: lot_parts={n:saved.read(n) for n in saved.namelist()}
                lot_meta={'schema':1,'groups':{},'rows':[{'row':i,'account':p['account'],'id':lot_key(p['account'],p['legal_description'],p.get('subdivision','')),'description':p['legal_description'],'area_source':p.get('area_source',{}),'unknown_area':p.get('land_area') is None} for i,p in enumerate(selected,2)]}
                put_metadata(lot_parts,lot_meta)
                with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as saved:
                    for n,value in lot_parts.items():saved.writestr(n,value)
            original=Path(application.get('source_pdf',''))
            if wizard.endswith('_renewal') and application.get('renewal_lookup'):
                from .designation_lookup import workbook_bytes
                updated,_,_=workbook_bytes(out.read_bytes(),application['renewal_lookup'])
                out.write_bytes(updated)
            if original.is_file(): shutil.copy2(original,sources/'Application.pdf')
            for p in selected:
                for index,record in enumerate(p.get('record_pdfs',[]),1):
                    if Path(record).is_file(): shutil.copy2(record,sources/f"HCAD_{p['account']}_{index}.pdf")
            audit={'schema':2,'wizard':wizard,'created_utc':datetime.now(timezone.utc).isoformat(),'fields':fields,
                'input_method':application.get('input_method','Application PDF'),
                'application_suggestions':application.get('fields',{}),'detected_wizard':application.get('detected_wizard'),
                'designation_sources':application.get('designation_sources',[]),
                'renewal_lookup':application.get('renewal_lookup',{}),
                'application_sha256':hashlib.sha256(original.read_bytes()).hexdigest() if original.is_file() else '',
                'template_sha256':hashlib.sha256(template.read_bytes()).hexdigest(),
                'sources':parcel_data.get('sources',[]),'warnings':parcel_data.get('warnings',[])+calculation_review_notes(fields,selected,wizard),
                'parcels':parcel_data['parcels'],'unresolved_records':parcel_data.get('unresolved_records',[]),
                'property_hints':application.get('property_hints',[])}
            if reviewed:
                audit.update(planner_accepted_selection=True, submitted_fields=submitted_fields, submitted_parcels=submitted_parcels)
            (sources/'Import review.json').write_text(json.dumps(audit,indent=2),encoding='utf8')
            if audit['warnings']:
                (sources/'Review notes.txt').write_text('Calculation sheet review notes\n\n'+'\n\n'.join(audit['warnings']),encoding='utf8')
            # Stage all supporting files before committing the new workbook.
            shutil.copytree(sources,source_dir)
            try:
                with destination.open('xb') as f: f.write(out.read_bytes())
            except Exception:
                # Only the newly-created source folder, never a pre-existing case.
                if source_dir.resolve().parent==destination.parent: shutil.rmtree(source_dir)
                raise
    return destination
