"""Building-line PC facts, with area and developed-lot denominators kept separate."""
import io
import re
from decimal import Decimal
from openpyxl import load_workbook
from core.boundary_map import clean, fingerprint, format_block_range
from wizards.mls.pc_data import ELIGIBLE, number, percent, date_text, normalized, read_shared_bytes


def read_pc_facts(path, source_path=None, wizard_key='mbl'):
    renewal = wizard_key == 'mbl_renewal'
    wb = load_workbook(io.BytesIO(read_shared_bytes(path)), data_only=True)
    try:
        ws = next(s for s in wb if s.title.lower().startswith('step 2'))
        title = normalized(ws['A1'].value)
        if 'BUILDING LINE' not in title or ('RENEWAL' in title) != renewal:
            raise ValueError('Select an MBL'+(' Renewal' if renewal else '')+' calculation sheet for these PC documents.')
        ci = {clean(a):date_text(b) for a,b in wb['Case Info'].iter_rows(max_col=2,values_only=True) if clean(a)} if 'Case Info' in wb else {}
        ci['Original_Passage_Date'] = ci.get('Original_Passage_Date') or ci.get('Original_Adoption_Date') or ci.get('Original_Council_Passage_Date','')
        ci['PC_Agenda_Number'] = ci.get('PC_Agenda_Number') or ci.get('PC_Agenda_Item','')
        modern = 'BUILDING LINE TO RENEW' in normalized(ws['C2'].value) or (number(ws['D2'].value) or 0)>0
        minimum = number(ws['D2' if modern else 'E10'].value) if renewal else number(ws['L16'].value)
        if not renewal and (minimum is None or minimum<=0):
            match=re.fullmatch(r'\s*([\d,.]+)\s*(?:-?\s*(?:FOOT|FEET|FT)\.?)?\s*',normalized(ws['E4'].value))
            minimum=number(match[1]) if match else None
        if minimum is None or minimum<=0:
            raise ValueError('Save a positive building-line minimum in Step 2 before making the PC documents. The wizard does not guess a minimum from the ranked lots.')
        header=next((r for r in range(12,21) if normalized(ws.cell(r,1).value) in ('PROPERTY DESCRIPTION','ADDRESS')),None)
        if header is None:raise ValueError('The Step 2 property table could not be identified.')
        columns={normalized(c.value):c.column for c in ws[header] if clean(c.value)}
        support_col=next((v for k,v in columns.items() if k in ('SIGNED','SUPPORT','SIGNED IN SUPPORT','SIGNED PETITION')),None)
        step1=next((s for s in wb if s.title.lower().startswith('step 1')),None)
        groups={}
        if step1 is not None:
            cols={normalized(c.value):c.column-1 for c in step1[1] if clean(c.value)}
            if 'PROPERTY DESCRIPTION' in cols and 'GROUP_DSCR' in cols:
                for values in step1.iter_rows(min_row=2,values_only=True):
                    desc=normalized(values[cols['PROPERTY DESCRIPTION']]);group=clean(values[cols['GROUP_DSCR']])
                    if desc and group and group.upper()!='<NULL>':groups.setdefault(desc,set()).add(group)
        rows=[];warnings=[]
        aliases={'MULTIFAMILY':'MF','MULTI FAMILY':'MF','COMMERCIAL':'COM','VACANT':'VAC','EXCLUDED':'EXC'}
        for r in range(header+1,ws.max_row+1):
            desc=clean(ws.cell(r,1).value)
            if not desc or desc in ('0','0.0'):continue
            lu=normalized(ws.cell(r,2).value);lu=aliases.get(lu,lu)
            area=number(ws.cell(r,5 if renewal else 4).value)
            bl=number(ws.cell(r,4 if renewal else 5).value)
            support=normalized(ws.cell(r,support_col).value) if support_col else ''
            known=bool(support_col) and support in ('','0','N','NO','Y','YES','1','TRUE','FALSE')
            if lu not in ELIGIBLE|{'EXC'} or area is None or area<=0:
                warnings.append(f'Step 2 row {r}: review land use and lot area.')
            if lu in ELIGIBLE-{'VAC'} and (bl is None or bl<0):
                warnings.append(f'Step 2 row {r}: enter the current building line in feet; an empty measurement is not zero.')
            if support_col and not known:warnings.append(f'Step 2 row {r}: review the support entry.')
            g=groups.get(normalized(desc),set())
            rows.append(dict(row=r,description=desc,land_use=lu,area=float(area) if area is not None else None,
                building_line=float(bl) if bl is not None else None,map_land_use=next(iter(g)) if len(g)==1 else '',
                signed=known and support in ('Y','YES','1','TRUE'),support_known=known,
                renewal_flag=normalized(ws.cell(r,6).value) if renewal else ''))
        if not rows:raise ValueError('Step 2 has no saved property rows. Save the calculated workbook in Excel first.')
        eligible=[r for r in rows if r['land_use'] in ELIGIBLE]
        complete=all(r['land_use'] in ELIGIBLE|{'EXC'} and r['area'] is not None and r['area']>0 for r in rows)
        total=sum(Decimal(str(r['area'])) for r in eligible) if complete else None
        sfr=[r for r in eligible if r['land_use']=='SFR'];signed=[r for r in eligible if r['signed']]
        developed=[r for r in eligible if r['land_use']!='VAC']
        bl_complete=all(r['land_use'] in ELIGIBLE|{'EXC'} for r in rows) and all(r['building_line'] is not None and r['building_line']>=0 for r in developed)
        qualifying=[r for r in developed if r['building_line'] is not None and Decimal(str(r['building_line']))>=minimum]
        support_complete=complete and all(r['support_known'] for r in eligible)
        sides=list(dict.fromkeys(re.sub(r'\s+sides?$','',clean(ws[c].value),flags=re.I).lower() for c in ('B6','D6') if clean(ws[c].value)))
        side_text=' and '.join(sides)+(' sides' if len(sides)>1 else ' side')
        street=clean(ws['B4'].value);block=format_block_range(ws['D4'].value)
        if street and normalized(street) not in normalized(block):block+=' of '+street
        cross=[clean(ws[c].value).rstrip(',.') for c in ('B5','D5')]
        boundary=re.sub(r'\b(\d+)(st|nd|rd|th)\b',lambda m:m[1]+m[2].lower(),f'{block}, {side_text}, between {cross[0]} and {cross[1]}',flags=re.I)
        app=clean(ws['B2'].value)
        if not re.fullmatch(r'\d+(?:\s*REN)?' if renewal else r'\d+',app,flags=re.I):raise ValueError('Enter the application number in Step 2 B2.')
        if renewal:app=re.sub(r'\s*REN$','',app,flags=re.I)+'REN'
        if not street or not block or not all(cross) or not sides:warnings.append('Complete the boundary streets, block range and sides in Step 2.')
        if renewal:
            warnings.append('The SharePoint renewal report retains 60% land-area and 70% structure criteria, while the renewal calculation uses different tests. Confirm the applicable renewal criteria before presenting; figures below describe the saved rows without asserting eligibility.')
        return dict(wizard_key=wizard_key,workbook_hash=fingerprint(source_path or path),application=app,
            minimum=float(minimum),minimum_text=format(minimum,'f').rstrip('0').rstrip('.') if '.' in format(minimum,'f') else str(minimum),
            boundary=boundary,block=block,street=street,sides=side_text,blockfaces=len(sides),cross_streets=cross,rows=rows,
            all_count=len(rows),eligible_count=len(eligible),excluded_count=sum(r['land_use']=='EXC' for r in rows),
            sfr_count=len(sfr),sfr_lot_percent=percent(len(sfr),len(eligible)) if complete else '[review]',
            sfr_area_percent=percent(sum(Decimal(str(r['area'])) for r in sfr),total) if complete else '[review]',
            signed_count=len(signed) if support_complete else None,
            support_area_percent=percent(sum(Decimal(str(r['area'])) for r in signed),total) if support_complete else '[not recorded]',
            eligible_area=float(total) if total else None,qualifying_count=len(qualifying) if bl_complete else None,
            qualifying_area_percent=percent(sum(Decimal(str(r['area'])) for r in qualifying),total) if complete and bl_complete else '[review]',
            developed_count=len(developed),developed_percent=percent(len(qualifying),len(developed)) if bl_complete else '[review]',
            below_count=len(developed)-len(qualifying) if bl_complete else None,vacant_count=sum(r['land_use']=='VAC' for r in eligible),
            case_info=ci,warnings=warnings)
    finally:wb.close()
