"""PC map categories; calculation eligibility and map land use are separate."""
from __future__ import annotations
import copy
import re
from core.boundary_map import CITY_PARCEL_SERVICE, parcel_polygon

# Sampled from the planner's PC legend. Keep one palette for lots and legend.
COLORS = {'SFR': '#FFFAD5', 'MF': '#FED362', 'COM': '#D9797A',
          'OFF': '#FD8AFF', 'IND': '#BAA9D9', 'PUB': '#ADDFF8',
          'TRN': '#FFFFFF', 'PRK': '#4ADA49', 'VAC': '#A5A5A5',
          'AGR': '#5D8944', 'UNK': '#FFFFFF'}
NAMES = {'SFR': 'Single-Family Residential', 'MF': 'Multi-Family Residential',
         'COM': 'Commercial', 'OFF': 'Office', 'IND': 'Industrial',
         'PUB': 'Public & Institutional', 'TRN': 'Transportation & Utility',
         'PRK': 'Park & Open Spaces', 'VAC': 'Undeveloped',
         'AGR': 'Agriculture Production', 'UNK': 'Unclassified / unavailable'}
LEGEND_USES = tuple(k for k in NAMES if k != 'UNK')
TRANSPORT_HATCH = '#55C8B9'
BELOW_HATCH = '#777777'
SUPPORT_COLOR = '#0066FF'
PROTEST_COLOR = '#FF0000'
LANDUSE_FIELDS = 'OBJECTID,HCAD_NUM,Tax_Year,group_dscr,group_cd,landuse_dscr'

def normalized(value):
    return re.sub(r'[^A-Z0-9]', '', str(value or '').upper())

def category(value):
    key = normalized(value)
    for code, name in NAMES.items():
        if key in (code, normalized(name)):
            return code
    return {'MULTIFAMILY':'MF', 'VACANT':'VAC', 'PARKANDOPENSPACES':'PRK',
            'PUBLICANDINSTITUTIONAL':'PUB', 'TRANSPORTATIONANDUTILITY':'TRN'}.get(key, 'UNK')

def context_use(feature):
    info = feature.get('pc_land_use', {})
    if info.get('category') in COLORS:
        return info['category']
    attrs = {k.lower():v for k,v in feature.get('attributes', {}).items()}
    return category(attrs.get('group_dscr'))

def selected_use(match):
    detailed = category(match.get('pc', {}).get('map_land_use'))
    if detailed != 'UNK':
        return detailed
    public = context_use(match['feature'])
    if public != 'UNK':
        return public
    # EXC is eligibility, not a land use. Never paint it as undeveloped.
    return category(match.get('pc', {}).get('land_use', match['property'].get('land_use')))

def attach_land_use(data, client):
    """Join detailed City uses by account and location; do not change geometry."""
    result = copy.deepcopy(data)
    client.progress('Loading detailed land-use colors for the PC map...')
    try:
        public = client.query(CITY_PARCEL_SERVICE, extent=data['extent'], fields=LANDUSE_FIELDS)
    except Exception as error:
        client.check_cancelled()
        result['warnings'].append('Detailed City land-use data could not be loaded. Calculation-sheet land uses are still shown; unavailable surrounding uses remain white. '+str(error))
        return result
    by_account = {}
    for feature in public:
        account = str(feature['attributes'].get('HCAD_NUM') or '').strip()
        if account:
            by_account.setdefault(account, []).append(feature)
    missing = 0
    years = set()
    for feature in [*result['context'], *(m['feature'] for m in result['matched'])]:
        account = str(feature['attributes'].get('HCAD_NUM') or '').strip()
        candidates = by_account.get(account, [])
        valid = []
        for candidate in candidates:
            try:
                a, b = parcel_polygon(feature), parcel_polygon(candidate)
                if a.intersection(b).area >= .9 * max(a.area, b.area):
                    valid.append(candidate)
            except Exception:
                continue
        uses = {category(f['attributes'].get('group_dscr')) for f in valid}
        if len(uses) == 1 and 'UNK' not in uses:
            attrs = valid[0]['attributes']
            feature['pc_land_use'] = {'category': uses.pop(), 'source': CITY_PARCEL_SERVICE,
                                      'tax_year': attrs.get('Tax_Year'), 'account': account}
            years.update(str(f['attributes'].get('Tax_Year')) for f in valid if f['attributes'].get('Tax_Year'))
        else:
            missing += 1
    result['sources'] = list(client.sources)
    result['pc_land_use_years'] = sorted(years)
    if missing:
        result['warnings'].append('Some parcels have no unique detailed City land-use match. Application parcels use available worksheet classifications; unclassified parcels remain white.')
    if years:
        result['warnings'].append('Surrounding land-use colors use City parcel tax year(s) '+', '.join(sorted(years))+'. Application properties use the worksheet detailed land use when present.')
    return result
