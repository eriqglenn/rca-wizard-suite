"""Workbook-driven City boundary maps. Geometry uses EPSG:2278 survey feet.

No desktop GIS, owner mailing data, or City calculation reimplementation.
The workbook property list is authoritative; public services supply geometry.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
import hashlib
import io
import json
import math
import re
import threading
import uuid
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from openpyxl import load_workbook
from shapely.geometry import Polygon, MultiPolygon, LineString, box
from shapely.ops import unary_union
from .workbooks import read_shared_bytes

PARCEL_SERVICE = 'https://www.gis.hctx.net/arcgis/rest/services/HCAD/Parcels/MapServer/0'
CITY_PARCEL_SERVICE = 'https://city-gis.internal.example/arcgis/rest/services/DIR/LandbaseandRoads_gx/MapServer/6'
ROAD_SERVICE = 'https://city-gis.internal.example/arcgis/rest/services/DIR/LandbaseandRoads_gx/MapServer/9'
COUNTY_ROAD_SERVICE = 'https://www.gis.hctx.net/arcgis/rest/services/ITC/CTS_roads/MapServer/0'
COUNTY_ROAD_FIELDS = 'OBJECTID,NAME,TYPE,DIR_PREFIX,DIR_SUFFIX,STREET,ALIAS_1,ALIAS_2,ALIAS_3'
PARCEL_FIELDS = ('OBJECTID,HCAD_NUM,site_str_num,site_str_name,site_str_pfx,site_str_sfx,'
                 'site_zip,tax_year,land_sqft,state_class,land_use,legal_dscr_1,legal_dscr_2')
CITY_PARCEL_FIELDS = ('OBJECTID,HCAD_NUM,Str_num,Str_pfx,Str_name,Str_sfx,Tax_Year,'
                      'Total_Land_Area,Legal_Dscr_1,Legal_Dscr_2')
ROAD_FIELDS = 'OBJECTID,NAME,FULL_NAME,PREFIX,ST_TYPE,SUFFIX'
MAP_SIGNATURE = 'RCA_BOUNDARY_MAP_V5'
MAP_TYPES = {
    'mls': ('Special Minimum Lot Size', 'MLS', 'Square Feet'),
    'mbl': ('Special Minimum Building Line', 'MBL', 'Feet'),
    'mls_renewal': ('Special Minimum Lot Size Renewal', 'MLS', 'Square Feet'),
    'mbl_renewal': ('Special Minimum Building Line Renewal', 'MBL', 'Feet'),
}
ROAD_WORDS = {'ST': 'STREET', 'DR': 'DRIVE', 'RD': 'ROAD', 'AVE': 'AVENUE',
              'BLVD': 'BOULEVARD', 'LN': 'LANE', 'CT': 'COURT', 'PL': 'PLACE',
              'CIR': 'CIRCLE', 'PKWY': 'PARKWAY', 'TER': 'TERRACE', 'TRL': 'TRAIL',
              'N': 'NORTH', 'S': 'SOUTH', 'E': 'EAST', 'W': 'WEST'}


class MapError(ValueError):
    """An actionable case, matching or layout error shown in the preview."""


def clean(value):
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return re.sub(r'\s+', ' ', str(value)).strip()


def road_name(value):
    return ' '.join(ROAD_WORDS.get(w, w) for w in re.findall(r'[A-Z0-9]+', clean(value).upper()))


def road_key(value):
    # City centerlines use ordinal names (22ND); older sheets may say 22.
    return re.sub(r'\b(\d+)(?:ST|ND|RD|TH)\b',r'\1',road_name(value)).replace(' ', '')


def format_block_range(value):
    """Use a hyphen between block-range endpoints without changing street names."""
    number = r'(?:\d{1,3}(?:,\d{3})+|\d+)'
    pattern = (rf'^(?P<prefix>(?:the\s+)?)(?P<first>{number})\s*'
               rf'(?:blocks?\s+)?(?:and\b|&|to\b|through\b|[-\u2013\u2014])\s*'
               rf'(?P<last>{number})\s+(?P<unit>blocks?)\b')
    return re.sub(pattern,lambda m:f'{m["prefix"]}{m["first"]}-{m["last"]} {m["unit"]}',
                  clean(value).rstrip(',.'),count=1,flags=re.I)


def account_number(value):
    s = clean(value)
    if not s:
        return ''
    # HCAD identifiers are fixed-width text. Decimal conversion used to strip
    # both zeros from valid accounts such as 0071050000002 before validation.
    if s.isascii() and s.isdigit() and len(s) == 13 and int(s) > 0:
        return s
    try:
        n = Decimal(s)
        if n == n.to_integral_value() and n > 0:
            s = format(n, 'f').split('.')[0]
    except InvalidOperation:
        pass
    if not s.isascii() or not s.isdigit() or len(s) not in (11, 12, 13) or int(s) <= 0:
        raise MapError(f'Invalid HCAD account {s!r}; use the 13-digit account, including leading zeros.')
    return s.zfill(13)


def fingerprint(path):
    return hashlib.sha256(read_shared_bytes(path)).hexdigest()


def minimum_number(value):
    # Read a calculated value/label; never calculate a minimum from property areas.
    m = re.fullmatch(r'\s*([\d,]+(?:\.\d+)?)\s*(?:-?\s*(?:Square\s*Foot|Square\s*Feet|Sq\.?\s*Ft\.?|Foot|Feet))?\s*',
                     clean(value), re.I)
    if not m:
        raise MapError('The minimum value is blank or invalid in Step 2. Save and recalculate the sheet in Excel, then retry.')
    n = Decimal(m.group(1).replace(',', ''))
    if n <= 0:
        raise MapError('The Step 2 minimum must be greater than zero.')
    return format(n, ',f').rstrip('0').rstrip('.') if '.' in format(n, ',f') else format(n, ',f')


@dataclass
class MapProperty:
    row: int
    account: str
    address: str
    number: str
    street: str
    land_use: str
    label: str
    legal_description: str = ''
    subdivision: str = ''
    lot_rows: list = field(default_factory=list)


@dataclass
class MapCase:
    wizard: str
    reference: str
    title: str
    block: str
    sides: str
    street: str
    cross_streets: list[str]
    minimum: str
    minimum_cell: str
    properties: list[MapProperty]
    workbook_hash: str
    warnings: list[str] = field(default_factory=list)
    lot_metadata: dict = field(default_factory=dict)

    @property
    def boundary_legend(self):
        return MAP_TYPES[self.wizard][0].replace(' Renewal', '') + ' Boundary'


def read_map_case(path, wizard, source_path=None):
    """Read the same Step 2 result cells as the wizards, after their cache resolver.

    Accept both established Step 2 sheet names. Do not read the owner mail-merge
    list: it may repeat or omit parcels according to ownership/notification rules.
    """
    if wizard not in MAP_TYPES:
        raise MapError('Unknown wizard type.')
    wb = load_workbook(io.BytesIO(read_shared_bytes(path)), data_only=True, read_only=True)
    try:
        s1 = next((s for s in wb if s.title.lower().startswith('step 1')), None)
        s2 = next((s for s in wb if s.title.lower().startswith('step 2')), None)
        if s1 is None or s2 is None:
            raise MapError('Select a calculation sheet with Step 1 and Step 2 tabs.')
        title, prefix, units = MAP_TYPES[wizard]
        heading = clean(s2['A1'].value).upper()
        if ('BUILDING LINE' not in heading if prefix == 'MBL' else 'LOT SIZE' not in heading) or \
                (('RENEWAL' in heading) != wizard.endswith('_renewal')):
            raise MapError('This calculation sheet belongs to a different wizard type. Select the matching workbook.')
        app = clean(s2['B2'].value)
        if not re.fullmatch(r'\d+(?:\s*REN)?', app, re.I):
            raise MapError('Enter the application number in Step 2 B2.')
        app = re.sub(r'\s*REN$', '', app, flags=re.I)
        reference = f'{prefix} {app}' + ('REN' if wizard.endswith('_renewal') else '')
        if wizard.endswith('_renewal'):
            # 2026 renewals use D2; the established older layout uses E10.
            old = 'most frequent' in clean(s2['D10'].value).lower() and not clean(s2['D2'].value)
            minimum_cell = 'E10' if old else 'D2'
        else:
            minimum_cell = 'L16' if wizard == 'mbl' and clean(s2['L16'].value) else 'E4'
        minimum = minimum_number(s2[minimum_cell].value) + ' ' + units
        block = format_block_range(s2['D4'].value)
        street = clean(s2['B4'].value)
        cross = [clean(s2[c].value).rstrip(',.') for c in ('B5', 'D5')]
        sides = [re.sub(r'\s+sides?$', '', clean(s2[c].value), flags=re.I).lower() for c in ('B6', 'D6')]
        sides = list(dict.fromkeys(s for s in sides if s))
        if not block or not street or not all(cross) or not sides:
            raise MapError('Complete the street, block, cross streets and side(s) in Step 2 B4/D4, B5/D5 and B6/D6.')
        sides_text = ' and '.join(sides) + (' sides' if len(sides) > 1 else ' side')
        header = {clean(c.value).lower(): c.column for c in next(s1.iter_rows(min_row=1, max_row=1))}
        def cell(row, *names):
            for name in names:
                col = header.get(name.lower())
                v = clean(row[col - 1].value) if col else ''
                if v and v != '0' and not v.startswith('#'):
                    return v
            return ''
        # Step 2 land-use entries can be planner overrides of the mirrored Step 1 values.
        table_header = next((r for r in range(12, 20) if clean(s2.cell(r, 1).value).lower() == 'property description'), None)
        if table_header is None:
            raise MapError('Could not locate the Step 2 property table.')
        land_by_desc = {}
        for row in s2.iter_rows(min_row=table_header + 1, max_row=min(s2.max_row or 600, 600), max_col=2):
            desc, lu = (clean(c.value) for c in row)
            if desc and lu:
                land_by_desc.setdefault(desc.upper(), set()).add(lu.upper())
        from .lot_divisions import read_metadata
        lot_metadata=read_metadata(wb)
        reviewed_ids={r['id'] for r in lot_metadata.get('rows',[])}
        properties, warnings, seen = [], [], {}
        # The installed MBL instructional sheet mixes real Haver parcels with
        # fictitious street-name exercises. Demonstrate those real parcel rows
        # explicitly as an example; never override a planner's saved case sheet.
        from .workbooks import is_bundled_input
        example = (wizard == 'mbl' and is_bundled_input(source_path or path)
                   and app == '0' and road_key(street) == road_key('Golden Retriever Lane'))
        if example:
            street, block, cross = 'Haver Street', '1800 block of Haver Street', ['Windsor Street', 'Mandell Street']
            sides_text = 'north side'
            reference = 'MBL EXAMPLE'
            warnings.append('INSTRUCTIONAL EXAMPLE: the default workbook contains fictitious heading streets. This preview uses its five north-side Haver Street property rows between Windsor and Mandell. Select your saved case workbook for an application map.')
        for rn, row in enumerate(s1.iter_rows(min_row=2), start=2):
            number = cell(row, 'Str_num', 'site_str_num')
            raw_street = cell(row, 'Str_name', 'site_str_name')
            address = cell(row, 'Site_addr_1', 'site_address', 'address')
            account_raw = cell(row, 'HCAD_NUM_1', 'HCAD_NUM', 'ALT_TAX_ID', 'TAX_ID')
            if not any((number, raw_street, address, account_raw)):
                continue
            account = account_number(account_raw) if account_raw else ''
            full_street = ' '.join(x for x in (cell(row, 'Str_pfx'), raw_street,
                                               cell(row, 'Str_sfx'), cell(row, 'Str_sfx_dir')) if x)
            if not address:
                address = ' '.join(x for x in (number, cell(row, 'str_num_sfx'), full_street) if x)
            if not number and address:
                m = re.match(r'^(\d+)\s+(.+)', address)
                if m:
                    number = m.group(1)
                    full_street = full_street or m.group(2)
            if number == '0':
                number = ''
            if (not number or not full_street) and not account:
                raise MapError(f'Step 1 row {rn}: provide a property address or an HCAD account for an unaddressed parcel.')
            desc = cell(row, 'Property Description') or f"{number} ({cell(row, 'Legal_Dscr_1')})"
            lu_values = land_by_desc.get(desc.upper(), set())
            if len(lu_values) > 1:
                raise MapError(f'Step 1 row {rn}: conflicting land-use classifications in Step 2.')
            lu = next(iter(lu_values), '') or cell(row, 'Land Use').upper()
            lu = {'SINGLE FAMILY': 'SFR', 'MULTI FAMILY': 'MF', 'COMMERCIAL': 'COM', 'VACANT': 'VAC',
                  'EXCLUDED': 'EXC'}.get(lu, lu)
            if lu not in ('SFR', 'MF', 'COM', 'VAC', 'EXC'):
                raise MapError(f'Step 1 row {rn} ({address}): set its Land Use to SFR, MF, COM, VAC or EXC in the calculation sheet.')
            from .lot_divisions import lot_key,parse_lots
            lot_row={'id':lot_key(account,cell(row,'Legal_Dscr_1'),cell(row,'Legal_Dscr_2')),'row':rn,'account':account,'address':address,'number':number,'street':full_street,'label':number,'land_use':lu,'legal_description':cell(row,'Legal_Dscr_1'),'subdivision':cell(row,'Legal_Dscr_2')}
            key = account or road_key(address)
            if key in seen:
                prev = seen[key]
                distinct=lot_row['id'] not in {r['id'] for r in prev.lot_rows}
                if distinct and ((parse_lots(lot_row['legal_description']) and all(parse_lots(r['legal_description']) for r in prev.lot_rows)) or (lot_row['id'] in reviewed_ids and all(r['id'] in reviewed_ids for r in prev.lot_rows))):
                    prev.lot_rows.append(lot_row)
                    continue
                if road_key(prev.address) != road_key(address) or prev.land_use != lu:
                    raise MapError(f'Step 1 rows {prev.row} and {rn}: the same HCAD account has different addresses or land uses. Resolve the match before mapping.')
                warnings.append(f'Duplicate property rows {prev.row}/{rn} are drawn once: {address}.')
                continue
            p = MapProperty(rn, account, address, number, full_street, lu, number,
                            cell(row, 'Legal_Dscr_1'), cell(row, 'Legal_Dscr_2'))
            p.lot_rows=[lot_row]
            properties.append(p)
            seen[key] = p
        if not properties:
            raise MapError('Step 1 contains no property addresses to map.')
        if example and ({p.account for p in properties} != {'0382500020001','0382500020002','0382500020003','0382500020004','0382500020005'}
                        or any(road_key(p.street) != road_key('Haver Street') for p in properties)):
            raise MapError('The instructional parcel rows have changed. Save a case workbook and complete its actual Step 2 street description before mapping.')
        if not any(road_key(p.street) == road_key(street) for p in properties):
            raise MapError(f'The Step 2 heading says {street}, but none of the Step 1 properties use that street. Correct the calculation sheet before creating its map.')
        return MapCase(wizard, reference, title, block, sides_text, street, cross, minimum,
                       f'{s2.title}!{minimum_cell}', properties, fingerprint(source_path or path), warnings,
                       lot_metadata)
    finally:
        wb.close()


def parcel_polygon(feature):
    rings = [Polygon(r) for r in feature.get('geometry', {}).get('rings', [])]
    if not rings or any(not p.is_valid or p.area <= 0 for p in rings):
        raise MapError('HCAD returned an invalid parcel outline.')
    if len(rings) == 1:
        return rings[0]
    # Classify by containment depth, preserving holes, islands and separate parts.
    parents = []
    for i, ring in enumerate(rings):
        containers = [j for j, p in enumerate(rings) if i != j and p.area > ring.area and p.contains(ring)]
        parents.append(min(containers, key=lambda j: rings[j].area) if containers else None)
    def depth(i):
        return 0 if parents[i] is None else 1 + depth(parents[i])
    parts = [Polygon(p.exterior.coords, [rings[j].exterior.coords for j in range(len(rings)) if parents[j] == i])
             for i, p in enumerate(rings) if depth(i) % 2 == 0]
    result = MultiPolygon(parts) if len(parts) > 1 else parts[0]
    if not result.is_valid:
        raise MapError('HCAD returned overlapping or invalid multipart parcel outlines.')
    return result


class PublicMapClient:
    """Bounded, paginated public queries with explicit local cache reuse."""
    def __init__(self, cache_dir, progress=None, cancel=None, use_cache=True):
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.progress = progress or (lambda message: None)
        self.cancel = cancel or threading.Event()
        self.sources = []
        self.use_cache = use_cache

    def check_cancelled(self):
        if self.cancel.is_set():
            raise MapError('Map creation cancelled.')

    def request(self, service, params):
        self.check_cancelled()
        url = service + '/query?' + urlencode({'f': 'json', **params})
        req = Request(url, headers={'User-Agent': 'Houston-RCA-Boundary-Map/1.20'})
        try:
            with urlopen(req, timeout=45) as response:
                data = json.load(response)
        except Exception as e:
            raise MapError(f'Could not reach the public map service. Check internet access and retry.\n{e}') from e
        if data.get('error'):
            raise MapError('Public map service: ' + clean(data['error'].get('message', data['error'])))
        self.check_cancelled()
        return data

    def query(self, service, where='1=1', extent=None, fields=PARCEL_FIELDS):
        params = {'where': where}
        if extent is not None:
            params.update(geometry=','.join(map(str, extent)), geometryType='esriGeometryEnvelope',
                          inSR=2278, spatialRel='esriSpatialRelIntersects')
        cache_key = hashlib.sha256(json.dumps([service, params, fields], sort_keys=True).encode()).hexdigest()
        cache_file = self.cache_dir / (cache_key + '.json')
        if self.use_cache and cache_file.exists():
            try:
                data = json.loads(cache_file.read_text(encoding='utf-8'))
                if data.get('cache_schema') == 1 and isinstance(data.get('features'), list):
                    self.sources.append({**data['_source'], 'cached': True})
                    return data['features']
            except (OSError, ValueError, KeyError):
                pass
        ids_result = self.request(service, {**params, 'returnIdsOnly': 'true'})
        ids = ids_result.get('objectIds') or []
        if len(ids) > 6000:
            raise MapError('This extent contains more than 6,000 parcels/road segments. Check for an outlying property address.')
        if len(ids) != len(set(ids)):
            raise MapError('The map service returned duplicate object IDs.')
        oid_field = ids_result.get('objectIdFieldName') or 'OBJECTID'
        features = []
        for offset in range(0, len(ids), 150):
            chunk = ids[offset:offset + 150]
            data = self.request(service, {'objectIds': ','.join(map(str, chunk)), 'outFields': fields,
                                         'returnGeometry': 'true', 'outSR': 2278})
            sr = data.get('spatialReference', {})
            if (sr.get('latestWkid') or sr.get('wkid')) not in (2278, 102740):
                raise MapError('The map service returned an unexpected coordinate system.')
            got = data.get('features', [])
            if data.get('exceededTransferLimit') or {f['attributes'].get(oid_field) for f in got} != set(chunk) or len(got) != len(chunk):
                raise MapError('The map service returned an incomplete result. Retry before using this map.')
            features.extend(got)
        source = {'service': service, 'query': params, 'retrieved_utc': datetime.now(timezone.utc).isoformat()}
        data = {'cache_schema': 1, '_source': source, 'features': features}
        temp = cache_file.with_suffix('.' + uuid.uuid4().hex + '.tmp')
        try:
            temp.write_text(json.dumps(data), encoding='utf-8')
            temp.replace(cache_file)
        finally:
            temp.unlink(missing_ok=True)
        self.sources.append({**source, 'cached': False})
        return features


def feature_address(f):
    a = f['attributes']
    return ' '.join(clean(a.get(k)) for k in ('site_str_num', 'site_str_pfx', 'site_str_name', 'site_str_sfx') if clean(a.get(k)))


def city_parcels(client, where='1=1', extent=None):
    mapping = {'Str_num': 'site_str_num', 'Str_pfx': 'site_str_pfx', 'Str_name': 'site_str_name',
               'Str_sfx': 'site_str_sfx', 'Tax_Year': 'tax_year', 'Total_Land_Area': 'land_sqft',
               'Legal_Dscr_1': 'legal_dscr_1', 'Legal_Dscr_2': 'legal_dscr_2'}
    result = []
    for raw in client.query(CITY_PARCEL_SERVICE, where, extent=extent, fields=CITY_PARCEL_FIELDS):
        attrs = {mapping.get(k, k): v for k, v in raw['attributes'].items()}
        result.append({**raw, 'attributes': attrs, '_parcel_service': CITY_PARCEL_SERVICE})
    return result


def street_identity(value):
    words = road_name(value).split()
    types = {'STREET','DRIVE','ROAD','AVENUE','BOULEVARD','LANE','COURT','PLACE','CIRCLE','PARKWAY','TERRACE','TRAIL'}
    if words and words[-1] in types:
        words.pop()
    return ' '.join(words)


def address_agrees(prop, feature):
    a = feature['attributes']
    street = ' '.join(clean(a.get(k)) for k in ('site_str_pfx','site_str_name','site_str_sfx') if clean(a.get(k)))
    expected, actual = road_name(prop.street), road_name(street)
    compatible = (road_key(expected) == road_key(actual) or
                  ((expected == street_identity(expected) or actual == street_identity(actual))
                   and road_key(street_identity(street)) == road_key(street_identity(prop.street))))
    return bool(prop.number and prop.street) and clean(a.get('site_str_num')) == prop.number and compatible


def legal_agrees(prop, feature):
    return bool(prop.legal_description and road_key(prop.legal_description) == road_key(feature['attributes'].get('legal_dscr_1')))


def address_candidates(prop, client):
    if not prop.number.isdigit() or not prop.street:
        return []
    root = street_identity(prop.street).split()
    if len(root) > 1 and root[0] in {'NORTH','SOUTH','EAST','WEST'}:
        root.pop(0)
    name = ' '.join(root).replace("'", "''")
    where = f"site_str_num = {int(prop.number)} AND UPPER(site_str_name) = '{name}'"
    candidates = [f for f in client.query(PARCEL_SERVICE, where) if address_agrees(prop, f)]
    if len(candidates) > 1:
        exact = [f for f in candidates if legal_agrees(prop, f)]
        if len(exact) == 1:
            return exact
    return candidates


def load_case_roads(case, client, extent, warnings):
    """Street names aid labeling; missing names must not discard matched parcels."""
    try:
        roads = list(client.query(ROAD_SERVICE, extent=extent, fields=ROAD_FIELDS))
    except MapError as e:
        client.check_cancelled()
        roads = []
        warnings.append('Warning: City street data is unavailable. ' + str(e))
    names = list(dict.fromkeys([case.street, *case.cross_streets]))
    keys = {road_key(f['attributes'].get('FULL_NAME')) for f in roads}
    missing = [name for name in names if road_key(name) not in keys]
    if missing:
        client.progress('Checking Harris County for missing street names...')
        try:
            county = client.query(COUNTY_ROAD_SERVICE, extent=extent, fields=COUNTY_ROAD_FIELDS)
        except MapError as e:
            client.check_cancelled()
            county = []
            warnings.append('Warning: the additional county street lookup is unavailable. ' + str(e))
        for name in missing:
            additions = []
            for f in county:
                a = f['attributes']
                full = clean(a.get('STREET')) or ' '.join(clean(a.get(k)) for k in ('DIR_PREFIX','NAME','TYPE','DIR_SUFFIX') if clean(a.get(k)))
                aliases = [full, *(clean(a.get(k)) for k in ('ALIAS_1','ALIAS_2','ALIAS_3'))]
                if not any(road_key(alias) == road_key(name) for alias in aliases if alias):
                    continue
                # Preserve the worksheet's named alias for a county-named road.
                additions.append({**f, 'attributes': {**a, 'FULL_NAME': full if road_key(full) == road_key(name) else name},
                                  '_road_service': COUNTY_ROAD_SERVICE})
            if additions:
                roads.extend(additions)
                warnings.append(f'Warning: {name} was not named in the City street layer. Its label uses Harris County street geometry. The Step 2 description is unchanged.')
            else:
                warnings.append(f'Warning: {name} from Step 2 could not be located in the public street layers. The map can still be saved; review and drag this street label to its correct location. The Step 2 description is unchanged.')
    return roads


def fetch_map_data(case, client):
    client.progress('Matching calculation-sheet properties to HCAD...')
    accounts = [p.account for p in case.properties if p.account]
    found = {}
    for i in range(0, len(accounts), 35):
        where = 'HCAD_NUM IN (' + ','.join("'" + a + "'" for a in accounts[i:i + 35]) + ')'
        for f in client.query(PARCEL_SERVICE, where):
            found.setdefault(clean(f['attributes'].get('HCAD_NUM')), []).append(f)
    missing_accounts = [a for a in accounts if not found.get(a)]
    for i in range(0, len(missing_accounts), 35):
        where = 'HCAD_NUM IN (' + ','.join("'" + a + "'" for a in missing_accounts[i:i + 35]) + ')'
        for f in city_parcels(client, where):
            found.setdefault(clean(f['attributes'].get('HCAD_NUM')), []).append(f)
    matched, issues, warnings = [], [], list(case.warnings)
    for prop in case.properties:
        client.check_cancelled()
        candidates = found.get(prop.account, []) if prop.account else []
        valid_account = [f for f in candidates if address_agrees(prop, f) or
                         ((not prop.number or not prop.street) and legal_agrees(prop, f))]
        method = 'account'
        old_address = feature_address(candidates[0]) if len(candidates) == 1 else ''
        if len(valid_account) == 1:
            candidates = valid_account
        else:
            candidates = address_candidates(prop, client)
            method = 'address and legal description' if prop.legal_description else 'address'
        if len(candidates) != 1:
            reason = f' Address/account mismatch: the account lists {old_address}.' if old_address else ''
            issues.append(f'Row {prop.row}: {prop.address or "unaddressed parcel"} [{prop.account or "no account"}] - {len(candidates)} matches; an exact parcel match is required.{reason}')
            continue
        f = candidates[0]
        if method != 'account':
            method = 'address and legal description' if legal_agrees(prop, f) else 'address'
        if method != 'account' and prop.account:
            warnings.append(f'Row {prop.row}: matched {prop.address} by {method} to HCAD {f["attributes"]["HCAD_NUM"]}. The sheet account {prop.account} is stale or belongs to {old_address or "a different parcel"}; the workbook was not changed.')
        if f.get('_parcel_service') == CITY_PARCEL_SERVICE:
            warnings.append(f'Row {prop.row}: account {prop.account} uses the exact parcel retained by Houston Map Viewer (tax year {f["attributes"].get("tax_year", "unknown")}).')
        if method == 'account' and (not prop.number or not prop.street):
            warnings.append(f'Row {prop.row}: unaddressed {prop.land_use} parcel matched by account and legal description; shown as 0 on the map.')
        elif road_key(feature_address(f)) != road_key(prop.address):
            warnings.append(f'Row {prop.row}: matched the base parcel address {feature_address(f)}; the sheet address includes additional unit information ({prop.address}).')
        try:
            parcel_polygon(f)
        except MapError as e:
            issues.append(f'Row {prop.row}: {prop.address} - {e}')
            continue
        matched.append({'property': asdict(prop), 'feature': f, 'match_method': method})
    if issues:
        raise MapError('Some calculation-sheet properties could not be matched:\n\n' + '\n'.join(issues[:35]))
    if any(m['feature'].get('_parcel_service') == CITY_PARCEL_SERVICE for m in matched):
        # A retired lot can overlap a newer combined lot. Use one retained City
        # parcel edition for the whole case when all exact accounts are present.
        city_found = {}
        current_accounts = [clean(m['feature']['attributes']['HCAD_NUM']) for m in matched]
        for i in range(0, len(current_accounts), 35):
            where = 'HCAD_NUM IN (' + ','.join("'" + a + "'" for a in current_accounts[i:i + 35]) + ')'
            for f in city_parcels(client, where):
                city_found.setdefault(clean(f['attributes']['HCAD_NUM']), []).append(f)
        coherent = []
        for m in matched:
            p = MapProperty(**m['property'])
            choices = [f for f in city_found.get(clean(m['feature']['attributes']['HCAD_NUM']), [])
                       if address_agrees(p, f) or legal_agrees(p, f)]
            if len(choices) != 1:
                break
            coherent.append({**m, 'feature': choices[0]})
        if len(coherent) == len(matched):
            matched = coherent
            warnings.append('The map uses one consistent Houston Map Viewer parcel edition for all worksheet accounts, including retained parcels. Review the recorded tax year; current county parcels may have been combined or changed.')
    polys = [parcel_polygon(m['feature']) for m in matched]
    united = unary_union(polys)
    if sum(p.area for p in polys) - united.area > 1:
        raise MapError('Selected HCAD parcel shapes overlap. Review stacked or duplicate accounts before mapping.')
    x0, y0, x1, y1 = united.bounds
    if max(x1 - x0, y1 - y0) > 12000:
        raise MapError('The properties extend over 12,000 feet. Check the calculation sheet for an outlying address.')
    pad = max(400, max(x1 - x0, y1 - y0) * .65)
    extent = [x0 - pad, y0 - pad, x1 + pad, y1 + pad]
    client.progress('Loading surrounding parcels and street names...')
    context = (city_parcels(client, extent=extent) if any(m['feature'].get('_parcel_service') == CITY_PARCEL_SERVICE for m in matched)
               else client.query(PARCEL_SERVICE, extent=extent))
    roads = load_case_roads(case, client, extent, warnings)
    valid_context = []
    for f in context:
        try:
            parcel_polygon(f)
            valid_context.append(f)
        except MapError:
            warnings.append('One surrounding parcel has invalid geometry and could not be shown: ' + feature_address(f))
    return {'case': asdict(case), 'matched': matched, 'context': valid_context, 'roads': roads,
            'extent': extent, 'sources': client.sources, 'warnings': warnings}


def map_case(data):
    d = dict(data['case'])
    d['properties'] = [MapProperty(**p) for p in d['properties']]
    return MapCase(**d)


def application_boundary(data, include_street=False):
    # Every property in the calculation sheet belongs in the map boundary.
    # EXC is a classification/label, not an instruction to remove its geometry.
    selected = unary_union([parcel_polygon(m['feature']) for m in data['matched']])
    if not include_street:
        return selected
    # Close only street-width gaps and leave unlisted tax parcels outside.
    closing = selected.buffer(45, join_style=2).buffer(-45, join_style=2).intersection(selected.convex_hull)
    result = selected.union(closing)
    accounts = {clean(m['feature']['attributes'].get('HCAD_NUM')) for m in data['matched']}
    blockers = [parcel_polygon(f) for f in data['context'] if clean(f['attributes'].get('HCAD_NUM')) not in accounts]
    if blockers:
        result = result.difference(unary_union(blockers))
    result = result.union(selected)
    if not result.is_valid or not result.buffer(.001).covers(selected):
        raise MapError('Could not construct a valid boundary across the street. Use Parcel edges in the preview.')
    return result


def assert_map_current(map_path, workbook_path, wizard):
    """Return a provenance warning; never require replacement of a chosen map.

    The historical name remains for callers in all four wizards. Only an
    unreadable PDF raises MapError; age and case differences are advisory.
    Embedded provenance still works if a generated PDF is copied alone.
    """
    if not map_path or not Path(map_path).is_file() or Path(map_path).suffix.lower() != '.pdf':
        return
    from pypdf import PdfReader
    try:
        reader = PdfReader(map_path)
        tag = str((reader.metadata or {}).get('/Keywords', '') or '')
    except Exception as e:
        raise MapError('Could not read the selected boundary-map PDF: ' + str(e)) from e
    continuation = ' You can continue generating documents with the selected map.'
    older_formats = {
        'V1': 'This map predates the EXC boundary correction.',
        'V2': 'This map predates the removal of land-use labels and placement of addresses inside their lots.',
        'V3': 'This map predates the updated block-range formatting and map layout checks.',
        'V4': 'This map predates the addition of the City of Houston seal.',
    }
    for version, detail in older_formats.items():
        if tag.startswith('RCA_BOUNDARY_MAP_' + version + '|'):
            return 'Warning: ' + detail + continuation
    if not tag.startswith(MAP_SIGNATURE + '|'):
        return  # Existing externally prepared maps remain supported.
    pieces = tag.split('|')
    try:
        mismatch = len(pieces) != 3 or pieces[1] != wizard or pieces[2] != fingerprint(workbook_path)
    except OSError:
        return 'Warning: The selected map could not be compared with the calculation sheet.' + continuation
    if mismatch:
        return 'Warning: The calculation sheet or wizard type has changed since this boundary map was made.' + continuation
