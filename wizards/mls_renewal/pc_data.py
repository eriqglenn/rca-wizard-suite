"""Renewal PC facts: current areas, established minimum, and explicit support only."""
import io
import re
from decimal import Decimal
from openpyxl import load_workbook
from core.boundary_map import clean, fingerprint, format_block_range
try:
    from core.workbooks import read_shared_bytes
except ImportError:  # V26work reads closed workbooks directly.
    from pathlib import Path
    def read_shared_bytes(path): return Path(path).read_bytes()
from wizards.mls.pc_data import ELIGIBLE, number, percent, date_text, normalized


def read_pc_facts(path, source_path=None):
    wb = load_workbook(io.BytesIO(read_shared_bytes(path)), data_only=True)
    try:
        ws = next(s for s in wb if s.title.lower().startswith('step 2'))
        if not all(word in normalized(ws['A1'].value) for word in ('LOT SIZE', 'RENEWAL')):
            raise ValueError('Select an MLS Renewal calculation sheet for the renewal PC documents.')
        ci = {clean(a): date_text(b) for a, b in wb['Case Info'].iter_rows(max_col=2, values_only=True) if clean(a)} if 'Case Info' in wb else {}
        if not ci.get('Original_Passage_Date'):
            ci['Original_Passage_Date'] = ci.get('Original_Adoption_Date') or ci.get('Original_Council_Passage_Date') or ''
        if not ci.get('PC_Agenda_Number'):ci['PC_Agenda_Number'] = ci.get('PC_Agenda_Item', '')
        modern = 'LOT SIZE TO RENEW' in normalized(ws['C2'].value) or (number(ws['D2'].value) or 0) > 0
        minimum = number(ws['D2' if modern else 'E10'].value)
        if minimum is None or minimum <= 0:
            raise ValueError('Save a positive lot size to renew in Step 2 before making the PC documents.')
        header = next((r for r in range(12, 21) if normalized(ws.cell(r, 1).value) in ('PROPERTY DESCRIPTION', 'ADDRESS')), 14)
        columns = {normalized(c.value): c.column for c in ws[header] if clean(c.value)}
        support_col = next((columns[k] for k in ('SIGNED', 'SUPPORT', 'SIGNED IN SUPPORT', 'SIGNED PETITION') if k in columns), None)
        corner_col = next((v for k, v in columns.items() if 'CORNER' in k), None)
        step1 = next(s for s in wb if s.title.lower().startswith('step 1'))
        h1 = {normalized(c.value): c.column - 1 for c in step1[1] if clean(c.value)}
        groups = {}
        if 'PROPERTY DESCRIPTION' in h1 and 'GROUP_DSCR' in h1:
            for values in step1.iter_rows(min_row=2, values_only=True):
                desc = normalized(values[h1['PROPERTY DESCRIPTION']]); group = clean(values[h1['GROUP_DSCR']])
                if desc and group and group.upper() != '<NULL>': groups.setdefault(desc, set()).add(group)
        rows = []; warnings = []
        aliases = {'MULTIFAMILY': 'MF', 'MULTI FAMILY': 'MF', 'COMMERCIAL': 'COM', 'VACANT': 'VAC', 'EXCLUDED': 'EXC'}
        for r in range(header + 1, ws.max_row + 1):
            desc = clean(ws.cell(r, 1).value)
            if not desc or desc in ('0', '0.0'): continue
            lu = normalized(ws.cell(r, 2).value); lu = aliases.get(lu, lu)
            area = number(ws.cell(r, 5 if modern else 4).value)
            if lu not in ELIGIBLE | {'EXC'} or area is None or area < 0:
                warnings.append(f'Step 2 row {r}: review the current lot area and land use.')
            support = normalized(ws.cell(r, support_col).value) if support_col else ''
            corner = normalized(ws.cell(r, corner_col).value) if corner_col else ''
            groups_for_row = groups.get(normalized(desc), set())
            rows.append(dict(row=r, description=desc, land_use=lu, area=float(area) if area is not None else None,
                             map_land_use=next(iter(groups_for_row)) if len(groups_for_row) == 1 else '',
                             signed=support in ('Y', 'YES', '1', 'TRUE'),
                             support_known=bool(support_col) and support in ('Y', 'YES', '1', 'TRUE', 'N', 'NO', '0', 'FALSE'),
                             corner=corner in ('Y', 'YES', '1', 'TRUE') if corner in ('Y', 'YES', '1', 'TRUE', 'N', 'NO', '0', 'FALSE') else None,
                             original_area=float(number(ws.cell(r, 4).value) or 0) if modern else None,
                             renewal_flag=normalized(ws.cell(r, 6).value) if modern else ''))
        if not rows: raise ValueError('Step 2 has no saved property rows. Save the calculation sheet in Excel first.')
        eligible = [r for r in rows if r['land_use'] in ELIGIBLE]
        complete = all(r['land_use'] in ELIGIBLE | {'EXC'} and r['area'] is not None and (r['area'] > 0 or r['land_use'] == 'EXC') for r in rows)
        total_area = sum(Decimal(str(r['area'])) for r in eligible) if complete else None
        sfr = [r for r in eligible if r['land_use'] == 'SFR']
        qualifying = [r for r in eligible if r['area'] is not None and Decimal(str(r['area'])) >= minimum]
        sides = list(dict.fromkeys(re.sub(r'\s+sides?$', '', clean(ws[c].value), flags=re.I).lower() for c in ('B6', 'D6') if clean(ws[c].value)))
        side_text = ' and '.join(sides) + (' sides' if len(sides) > 1 else ' side')
        street = clean(ws['B4'].value); block = format_block_range(ws['D4'].value)
        if street and normalized(street) not in normalized(block): block += ' of ' + street
        cross = [clean(ws[c].value).rstrip(',.') for c in ('B5', 'D5')]
        boundary = re.sub(r'\b(\d+)(st|nd|rd|th)\b', lambda m: m[1] + m[2].lower(), f'{block}, {side_text}, between {cross[0]} and {cross[1]}', flags=re.I)
        if not re.fullmatch(r'\d+(?:\s*REN)?', clean(ws['B2'].value), flags=re.I):
            raise ValueError('Enter the renewal application number in Step 2 B2.')
        app = re.sub(r'\s*REN$', '', clean(ws['B2'].value), flags=re.I) + 'REN'
        support_complete = complete and all(r['support_known'] for r in eligible)
        signed = [r for r in eligible if r['signed']]
        variance = None
        if complete and eligible and all(r['corner'] is not None for r in eligible):
            noncorner = [Decimal(str(r['area'])) for r in eligible if not r['corner']]
            if noncorner:
                average = sum(noncorner) / len(noncorner)
                within = sum(average * Decimal('.9') <= a <= average * Decimal('1.1') for a in noncorner)
                variance = dict(count=len(noncorner), within=within, average=f'{average:,.0f}', percent=percent(within, len(noncorner)))
        return dict(wizard_key='mls_renewal', workbook_hash=fingerprint(source_path or path), application=app,
                    minimum=float(minimum), minimum_text=f'{minimum:,.0f}', boundary=boundary, block=block, street=street,
                    sides=side_text, blockfaces=len(sides), cross_streets=cross, rows=rows, all_count=len(rows),
                    eligible_count=len(eligible), excluded_count=len(rows)-len(eligible), sfr_count=len(sfr),
                    sfr_lot_percent=percent(len(sfr),len(eligible)) if complete else '[review]',
                    sfr_area_percent=percent(sum(Decimal(str(r['area'])) for r in sfr),total_area) if complete else '[review]',
                    signed_count=len(signed) if support_complete else None,
                    support_area_percent=percent(sum(Decimal(str(r['area'])) for r in signed),total_area) if support_complete else '[not recorded]',
                    eligible_area=float(total_area) if total_area else None, qualifying_count=len(qualifying) if complete else None,
                    qualifying_area_percent=percent(sum(Decimal(str(r['area'])) for r in qualifying),total_area) if complete else '[review]',
                    case_info=ci, variance=variance, warnings=warnings)
    finally:
        wb.close()
