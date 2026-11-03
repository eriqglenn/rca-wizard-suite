"""Read-only facts for the MLS Planning Commission report and thematic map.

Keep the denominator with each metric: lot counts are not land-area shares.
The official minimum is read from Step 2, never inferred from a rank index.
"""
from __future__ import annotations
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import io
import json
import re
from openpyxl import load_workbook
from core.boundary_map import clean, format_block_range, minimum_number, fingerprint
try:
    from core.workbooks import read_shared_bytes
except ImportError:  # V26work reads closed workbooks directly.
    from pathlib import Path
    def read_shared_bytes(path): return Path(path).read_bytes()

ELIGIBLE = {'SFR', 'MF', 'COM', 'VAC'}

def number(value):
    try:
        n = Decimal(str(value).strip().replace(',', ''))
        return n if n.is_finite() else None
    except (ValueError, InvalidOperation):
        return None

def percent(part, total):
    if part is None or not total:
        return '[review]'
    return str((Decimal(part) * 100 / Decimal(total)).quantize(Decimal('1'), rounding=ROUND_HALF_UP))

def date_text(value):
    if isinstance(value, (date, datetime)):
        return f'{value:%B} {value.day}, {value:%Y}'
    text = clean(value)
    for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%B %d, %Y'):
        try:
            return date_text(datetime.strptime(text, fmt))
        except ValueError:
            pass
    return text

def normalized(value):
    return re.sub(r'\s+', ' ', clean(value)).upper()

def read_pc_facts(path, source_path=None, wizard_key="mls"):
    """Use the resolved workbook's Step 2 rows; do not deduplicate by owner."""
    if wizard_key in ("mbl", "mbl_renewal"):
        from wizards.mbl.pc_data import read_pc_facts as read_mbl
        return read_mbl(path, source_path, wizard_key)
    if wizard_key == "mls_renewal":
        from wizards.mls_renewal.pc_data import read_pc_facts as read_renewal
        return read_renewal(path, source_path)
    wb = load_workbook(io.BytesIO(read_shared_bytes(path)), data_only=True)
    warnings = []
    try:
        ws = next(s for s in wb if s.title.lower().startswith('step 2'))
        if 'LOT SIZE' not in clean(ws['A1'].value).upper() or 'RENEWAL' in clean(ws['A1'].value).upper():
            raise ValueError('PC Draft Reports are available for the regular MLS wizard. Select an MLS calculation sheet.')
        ci = {}
        if 'Case Info' in wb:
            ci = {clean(row[0].value): row[1].value for row in wb['Case Info'].iter_rows(min_col=1, max_col=2) if clean(row[0].value)}
        minimum = number(minimum_number(ws['E4'].value))
        header = next(r for r in range(12, 21) if normalized(ws.cell(r, 1).value) == 'PROPERTY DESCRIPTION')
        # Step 2 has broad eligibility buckets. Step 1 retains the detailed
        # Houston land-use categories needed by the PC legend (office, park...).
        step1 = next(s for s in wb if s.title.lower().startswith('step 1'))
        columns = {normalized(c.value): c.column for c in step1[1] if clean(c.value)}
        descriptions = {}
        if 'PROPERTY DESCRIPTION' in columns and 'GROUP_DSCR' in columns:
            for values in step1.iter_rows(min_row=2, values_only=True):
                desc = normalized(values[columns['PROPERTY DESCRIPTION']-1])
                group = clean(values[columns['GROUP_DSCR']-1])
                if desc and group and group.upper() != '<NULL>':
                    descriptions.setdefault(desc, set()).add(group)
        rows = []
        aliases = {'MULTIFAMILY': 'MF', 'MULTI FAMILY': 'MF', 'COMMERCIAL': 'COM', 'VACANT': 'VAC', 'EXCLUDED': 'EXC'}
        for r in range(header + 1, ws.max_row + 1):
            desc, lu, signed, area = [ws.cell(r, c).value for c in range(1, 5)]
            if not clean(desc) or clean(desc) == '0':
                continue
            lu = aliases.get(normalized(lu), normalized(lu))
            area = number(area)
            support = normalized(signed)
            if lu not in ELIGIBLE | {'EXC'} or area is None or area < 0:
                warnings.append(f'Step 2 row {r}: review land use and lot area; affected percentages remain marked for review.')
            if support not in ('', '0', 'N', 'NO', 'Y', 'YES', '1', 'TRUE', 'FALSE'):
                warnings.append(f'Step 2 row {r}: support entry {clean(signed)!r} is not Y/N; support totals require review.')
            rows.append({'row': r, 'description': clean(desc), 'land_use': lu,
                         'map_land_use': next(iter(descriptions[normalized(desc)])) if len(descriptions.get(normalized(desc), ())) == 1 else '',
                         'area': float(area) if area is not None else None,
                         'signed': support in ('Y', 'YES', '1', 'TRUE'),
                         'support_known': support in ('', '0', 'N', 'NO', 'Y', 'YES', '1', 'TRUE', 'FALSE')})
        if not rows:
            raise ValueError('Step 2 has no property rows to analyze.')
        eligible = [r for r in rows if r['land_use'] in ELIGIBLE]
        excluded = [r for r in rows if r['land_use'] == 'EXC']
        complete = all(r['land_use'] in ELIGIBLE | {'EXC'} and r['area'] is not None and r['area'] >= 0 for r in rows)
        total_area = sum(Decimal(str(r['area'])) for r in eligible) if complete else None
        signed = [r for r in eligible if r['signed']]
        support_complete = complete and all(r['support_known'] for r in eligible)
        signed_area = sum(Decimal(str(r['area'])) for r in signed) if support_complete else None
        qualified = [r for r in eligible if r['area'] is not None and Decimal(str(r['area'])) >= minimum]
        qualified_area = sum(Decimal(str(r['area'])) for r in qualified) if complete else None
        sfr_count = sum(r['land_use'] == 'SFR' for r in eligible)
        sides = list(dict.fromkeys(re.sub(r'\s+sides?$', '', clean(ws[c].value), flags=re.I).lower() for c in ('B6', 'D6') if clean(ws[c].value)))
        block = format_block_range(ws['D4'].value)
        street = clean(ws['B4'].value)
        cross = [clean(ws[c].value).rstrip(',.') for c in ('B5', 'D5')]
        side_text = ' and '.join(sides) + (' sides' if len(sides) > 1 else ' side')
        boundary = f'{block}, {side_text}, between {cross[0]} and {cross[1]}'
        boundary = re.sub(r'\b(\d+)(st|nd|rd|th)\b', lambda m: m[1]+m[2].lower(), boundary, flags=re.I)
        result = {'workbook_hash': fingerprint(source_path or path), 'application': clean(ws['B2'].value),
                  'minimum': float(minimum), 'minimum_text': f'{minimum:,.0f}', 'boundary': boundary,
                  'block': block, 'street': street, 'sides': side_text, 'blockfaces': len(sides),
                  'cross_streets': cross, 'rows': rows, 'all_count': len(rows), 'eligible_count': len(eligible),
                  'excluded_count': len(excluded), 'sfr_count': sfr_count,
                  'sfr_lot_percent': percent(sfr_count, len(eligible)) if complete else '[review]',
                  'signed_count': len(signed) if support_complete else None,
                  'support_area_percent': percent(signed_area, total_area),
                  'eligible_area': float(total_area) if total_area is not None else None,
                  'qualifying_count': len(qualified) if complete else None,
                  'qualifying_area_percent': percent(qualified_area, total_area),
                  'case_info': {k: date_text(v) for k, v in ci.items()}, 'warnings': warnings}
        for cell, metric, name in [('F8', 'support_area_percent', 'support by area'), ('H11', 'sfr_lot_percent', 'SFR share by lot count')]:
            cached = number(ws[cell].value)
            actual = number(result[metric])
            if cached is not None and actual is not None and abs(cached*100-actual) > Decimal('0.51'):
                warnings.append(f'Step 2 {cell} differs from its property rows for {name}. Review the workbook; this draft uses the property rows.')
        return result
    finally:
        wb.close()

def attach_pc_rows(data, facts, protested_rows=()):
    """Join by full property description, never by sorted row position or owner."""
    import copy
    result = copy.deepcopy(data)
    by_desc = {}
    for row in facts['rows']:
        by_desc.setdefault(normalized(row['description']), []).append(row)
    warnings = list(facts['warnings'])
    protests = {int(r) for r in protested_rows}
    mapped_rows = set()
    for match in result['matched']:
        p = match['property']
        desc = f"{p.get('number') or '0'} ({p.get('legal_description', '')})"
        rows = by_desc.get(normalized(desc), [])
        # Some sheets use an address plus legal description without parentheses.
        if not rows:
            rows = [r for r in facts['rows'] if normalized(r['description']) == normalized(p.get('description', ''))]
        if len(rows) != 1:
            warnings.append(f"{p.get('address') or p.get('account')}: Step 2 description could not be uniquely linked; support and minimum symbols omitted.")
            match['pc'] = {'land_use': p['land_use'], 'signed': False, 'protested': False, 'below': False, 'known': False}
            continue
        row = rows[0]
        mapped_rows.add(row['row'])
        match['pc'] = {**row, 'protested': row['row'] in protests,
                       'below': (row['land_use'] in ELIGIBLE-{'VAC'} and row.get('building_line') is not None and 0 <= row['building_line'] < facts['minimum']) if facts.get('wizard_key','').startswith('mbl') else (row['land_use'] in ELIGIBLE and row['area'] is not None and row['area'] < facts['minimum']), 'known': True}
    for row in facts['rows']:
        if row['row'] not in mapped_rows:
            warnings.append(f"Step 2 row {row['row']} ({row['description']}): no separate matched parcel for this row. Review combined accounts or legal descriptions; its facts remain in the report totals.")
    result['pc_protested_rows']=list(protested_rows)
    result['pc_facts'] = facts
    result['warnings'] = list(dict.fromkeys(result.get('warnings', []) + warnings))
    return result

def preferences_path(source, output):
    import hashlib
    identity = hashlib.sha256(str(Path(source).resolve()).lower().encode()).hexdigest()[:16]
    return Path(output) / f'.mls_pc_{identity}.json'

def load_preferences(source, output, workbook_hash):
    try:
        saved = json.loads(preferences_path(source, output).read_text(encoding='utf8'))
        return saved.get('values', {}) if saved.get('workbook_hash') == workbook_hash else {}
    except (OSError, ValueError):
        return {}

def save_preferences(source, output, workbook_hash, values):
    import os, tempfile
    target = preferences_path(source, output)
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(dir=target.parent, prefix='.pc_', suffix='.tmp')
    try:
        with os.fdopen(fd, 'w', encoding='utf8') as stream:
            json.dump({'workbook_hash': workbook_hash, 'values': values}, stream, indent=2)
        os.replace(name, target)
    finally:
        Path(name).unlink(missing_ok=True)
