"""Fill the authoritative September 29 City templates without rewriting prose."""
from datetime import datetime,date
from decimal import Decimal,InvalidOperation,ROUND_HALF_UP
from pathlib import Path
import re
from docx import Document
from docx.oxml.ns import qn
from .property_letters import replace_span
from .mbl_templates_2026 import _checkboxes
from .routing_form import prepare_routing_form
from .docx_merge import _clear_all_red_run_colors
from .docx_unlock import unlock_docx

MARK='CITY_SHAREPOINT_20260929'
TOKEN=re.compile(r'\{\{CITY26_([A-Za-z0-9_]+)\}\}')
ELIGIBLE={'SFR','MF','COM','VAC'}
def number(value):
    try:
        n=Decimal(str(value).strip().replace(',','').replace('%',''))
        return n if n.is_finite() else None
    except (InvalidOperation,ValueError,TypeError):return None
def text(value):return '' if value is None else str(value).strip()
def percent(part,total):
    return str((Decimal(part)*100/Decimal(total)).quantize(Decimal('1'),rounding=ROUND_HALF_UP)) if total else '[VERIFY DENOMINATOR]'
def words(n):
    from wizards.mls.templates_2026 import _words
    return _words(int(n))
def date_text(value):
    if isinstance(value,(datetime,date)):return f'{value:%B} {value.day}, {value:%Y}'
    s=text(value)
    for fmt in ('%m/%d/%Y','%Y-%m-%d','%Y-%m-%d %H:%M:%S','%B %d, %Y'):
        try:return date_text(datetime.strptime(s,fmt))
        except ValueError:pass
    return s
def field_values(ctx,workbook,wizard):
    """Retain calculation contexts; distinguish area, count and adoption dates."""
    from openpyxl import load_workbook
    wb=load_workbook(workbook,data_only=True)
    try:
        ci={text(r[0]):r[1] for r in wb['Case Info'].iter_rows(min_col=1,max_col=2,values_only=True) if r[0]} if 'Case Info' in wb.sheetnames else {}
        rows=ctx.get('Property_Rows')
        if rows is None:
            # MLS mirrors the Step 1 rows into Step 2. Split legal lots remain
            # separate; never deduplicate these metrics by HCAD account/owner.
            s2=next(s for s in wb if s.title.lower().startswith('step 2'))
            rows=[]
            for row in s2.iter_rows(min_row=17,max_col=4,values_only=True):
                desc,lu,signed,area=row
                if text(desc) and text(lu).upper() in ELIGIBLE|{'EXC'}:
                    rows.append(dict(desc=text(desc),land_use=text(lu).upper(),signed=text(signed).upper() in ('Y','YES','1'),lot=number(area)))
            if not rows:
                s1=next(s for s in wb if s.title.lower().startswith('step 1'))
                headers={text(c.value).lower():c.column for c in s1[1]}
                def get(r,name):
                    c=headers.get(name.lower());return s1.cell(r,c).value if c else None
                for r in range(2,s1.max_row+1):
                    lu=text(get(r,'Land Use')).upper();area=number(get(r,'Lot Size'))
                    desc=text(get(r,'Property Description') or get(r,'Legal_Dscr_1'))
                    if desc and lu in ELIGIBLE|{'EXC'}:
                        signed=s2.cell(r+15,3).value
                        rows.append(dict(desc=desc,land_use=lu,signed=text(signed).upper() in ('Y','YES','1'),lot=area))
    finally:wb.close()
    renewal=wizard.endswith('_renewal');is_mbl=wizard.startswith('mbl')
    fields={k:text(v) for k,v in ctx.items() if not isinstance(v,(dict,list,tuple))}
    fields.update({k:date_text(v) for k,v in ci.items() if k in ('Original_Passage_Date','Original_Council_Passage_Date','Original_Adoption_Date','PC_Agenda_Item','PC_Presenter','PC_Motion_By','PC_Second_By','PC_Effective_Date','Council_Meeting_Date')})
    fields['NUMBER']=text(ctx.get('Application_Renewal_Number') if renewal else ctx.get('Application_Number'))
    fields['MINIMUM']=text(ctx.get('Minimum_Building_Line_Text') if is_mbl else ctx.get('Minimum_Lot_Size_Text'))
    fields['MINIMUM_WORDS']=text(ctx.get('Minimum_Building_Line_Words') if is_mbl else ctx.get('Minimum_Lot_Size_Words'))
    fields['AREA']=text(ctx.get('Area_Description'))
    fields['CAPTION']=text(ctx.get('Ordinance_Caption') or ctx.get('Ordinance_Area_Lower') or fields['AREA'])
    fields['DISTRICT']=text(ctx.get('Council_District'))
    fields['MEMBER']=text(ctx.get('Council_Member'))
    fields['PLANNER']=text(ctx.get('Planner_Name'))
    fields['PLANNER_CONTACT']=', '.join(text(ctx.get(k)) for k in ('Planner_Name','Planner_Position','Planner_Phone') if text(ctx.get(k)))
    fields['AUTHOR']=', '.join(text(ctx.get(k)) for k in ('Planner_Name','Planner_Phone') if text(ctx.get(k)))
    fields['APPLICANT_ADDRESS']=text(ctx.get('Applicant_Property_Address'))
    subdivision=text(ctx.get('Subdivision_Name'))
    fields['SUBDIVISION']=subdivision if 'subdivision' in subdivision.lower() else subdivision+' Subdivision'
    fields['ORDINANCE']=text(ctx.get('Original_Ordinance_No'))
    passage=next((ci.get(k) for k in ('Original_Passage_Date','Original_Council_Passage_Date','Original_Adoption_Date') if ci.get(k)),None)
    fields['PASSAGE_DATE']=date_text(passage) or '[Enter Original_Passage_Date in Case Info]'
    fields['PASSAGE_YEAR']=next(iter(re.findall(r'\b(?:19|20)\d{2}\b',date_text(passage))),'[Enter Original_Passage_Date in Case Info]')
    fields['EFFECTIVE_DATE']=date_text(ci.get('Original_Effective_Date')) or '[Enter Original_Effective_Date in Case Info]'
    fields['PC_DATE']=date_text(ci.get('PC_Approval_Date')) or 'Date'
    fields['DIRECTOR_DATE']=date_text(ci.get('Director_Approval_Date')) or 'Date'
    fields['DEADLINE']=date_text(ci.get('Council_Deadline')) or '[Enter Council_Deadline in Case Info]'
    fields['HEARING_DATE']=date_text(ci.get('PC_Hearing_Date')) or '[Enter PC_Hearing_Date in Case Info]'
    for key in ('PC_Agenda_Item','PC_Presenter','PC_Motion_By','PC_Second_By','PC_Effective_Date','Council_Meeting_Date'):
        fields[key]=date_text(ci.get(key)) or '________________'
    fields['PC_Agenda_Item']=text(ci.get('PC_Agenda_Item') or ci.get('PC_Agenda_Number')) or '________________'
    fields['PROTEST_SENTENCE']=text(ctx.get('Protest_Sentence'))
    n=number(ctx.get('Protest_Count'))
    if n is not None:
        fields['PROTEST_SENTENCE']='No protests were received.' if n==0 else f'{words(n).capitalize()} ({int(n)}) protest'+(' was received.' if n==1 else 's were received.')
    fields['TOTAL']=text(ctx.get('Total_Lot_Count'))
    fields['QUALIFYING']=text(ctx.get('Qualifying_Lot_Count'))
    fields['SUPPORT_AREA_PERCENT']=text(ctx.get('Support_Percentage_Number') or ctx.get('Support_Percentage')).rstrip('%')
    fields['SFR_LOT_PERCENT']=text(ctx.get('SFR_Percentage_Number') or ctx.get('SFR_Percentage')).rstrip('%')
    if wizard=='mls_renewal':rows=[dict(r,lot=r.get('cur')) for r in rows or []]
    eligible=[r for r in rows or [] if r.get('land_use') in ELIGIBLE]
    if eligible:
        fields['TOTAL']=str(len(eligible))
        sfr=[r for r in eligible if r['land_use']=='SFR']
        fields['SFR_LOT_PERCENT']=percent(len(sfr),len(eligible))
        areas=[number(r.get('lot')) for r in eligible]
        complete=all(a is not None and a>=0 for a in areas)
        total=sum(areas) if complete else None
        if total:
            fields['SFR_AREA_PERCENT']=percent(sum(number(r['lot']) for r in sfr),total)
            if not renewal:
                fields['SUPPORT_AREA_PERCENT']=percent(sum(number(r['lot']) for r in eligible if r.get('signed')),total)
        minimum=number(fields['MINIMUM'])
        if minimum is not None:
            dimension='cur' if wizard=='mbl_renewal' else 'bl' if is_mbl else 'lot'
            qualifying=[r for r in eligible if number(r.get(dimension)) is not None and number(r.get(dimension))>=minimum]
            fields['QUALIFYING']=str(len(qualifying))
            fields['QUALIFYING_AREA_PERCENT']=percent(sum(number(r['lot']) for r in qualifying),total) if total else '[VERIFY LOT AREAS]'
        fields['SFR_THRESHOLD']=len(sfr)*100>=len(eligible)*60
    else:
        fields['SFR_THRESHOLD']=None
    fields.setdefault('SFR_AREA_PERCENT','[VERIFY LOT AREAS]')
    fields.setdefault('QUALIFYING_AREA_PERCENT','[VERIFY LOT AREAS]')
    for name in ('TOTAL','QUALIFYING'):
        n=number(fields[name]);fields[name+'_PHRASE']=f'{words(n)} ({int(n)})' if n is not None else '[VERIFY LOT COUNT]'
    fields['QUALIFYING_OF_TOTAL']=fields['QUALIFYING']+' out of '+fields['TOTAL']
    fields['QUALIFYING_PHRASE_CAP']=fields['QUALIFYING_PHRASE'].capitalize()
    fields['MINIMUM_PHRASE']=fields['MINIMUM_WORDS']+' ('+fields['MINIMUM']+')'
    dev=number(ctx.get('Developed_Lot_Count'))
    fields['DEVELOPED_PHRASE']=f'{words(dev)} ({int(dev)})' if dev is not None else '[VERIFY DEVELOPED LOT COUNT]'
    noncorner=number(ci.get('PC_Noncorner_Lot_Count'))
    total_count=number(fields['TOTAL'])
    fields['NONCORNER_PHRASE']=f'{words(noncorner)} ({int(noncorner)})' if noncorner is not None and total_count is not None and noncorner==int(noncorner) and 0<=noncorner<=total_count else '[Enter PC_Noncorner_Lot_Count in Case Info]'
    fields['PLAT_YEAR']=text(ci.get('Subdivision_Plat_Year')) or '[Enter Subdivision_Plat_Year in Case Info]'
    fields['STRUCTURE_AGE']=text(ci.get('Structure_Age_Statement')) or '[Enter Structure_Age_Statement in Case Info]'
    return fields

def render_updated(template,output,ctx,workbook,wizard,role):
    """Return False for an older/custom template; retain its existing filler."""
    template=Path(template);doc=Document(template)
    if MARK not in (doc.core_properties.keywords or ''):return False
    is_pc=text(ctx.get('Approval_Path'))=='Planning Commission Approval' or (number(ctx.get('Protest_Count')) or 0)>0
    if role in ('routing','staff'):
        suffix='_PC' if is_pc else ''
        if role=='staff' and not wizard.endswith('_renewal'):
            support=number(ctx.get('Support_Percentage_Number') or ctx.get('Support_Percentage'))
            suffix='_Protest' if (number(ctx.get('Protest_Count')) or 0)>0 else '_Less51' if support is not None and support<51 else '_PC' if is_pc else ''
        alternative=template.with_name(template.stem+suffix+template.suffix) if suffix else template
        if role=='staff' and wizard=='mls_renewal' and is_pc and template.name=='MLS_REN_OneMeeting_Director_Approval.docx':
            alternative=template.with_name('MLS_REN_OneMeeting_PC_Approval.docx')
        if alternative!=template and alternative.exists():doc=Document(alternative)
    fields=field_values(ctx,workbook,wizard)
    for p in doc.element.iter(qn('w:p')):
        if list(p.iterdescendants(qn('w:p'))):continue
        content=''.join(t.text or '' for t in p.iter(qn('w:t')))
        for match in reversed(list(TOKEN.finditer(content))):
            key=match.group(1)
            if key not in fields:raise ValueError('Unsupported City template field: '+key)
            replace_span(p,match.start(),match.end(),text(fields[key]) or '[Enter '+key+' in Case Info]')
        content=''.join(t.text or '' for t in p.iter(qn('w:t')))
        if role=='routing':
            mbl=wizard.startswith('mbl');renewal=wizard.endswith('_renewal')
            if content.startswith('BUILDING LINE'):_checkboxes(p,[mbl,not mbl])
            elif content.replace('\u00a0',' ').startswith('Contains lots, at least 60'):_checkboxes(p,[fields['SFR_THRESHOLD'] is True,fields['SFR_THRESHOLD'] is False])
            elif 'SMLS Area Application' in content:_checkboxes(p,[False,not mbl and not renewal,mbl and not renewal])
            elif 'SMLS Block Renewal' in content:_checkboxes(p,[not mbl and renewal])
            elif 'SMBL Block Renewal' in content:_checkboxes(p,[mbl and renewal])
            elif "Director's Approval" in content and 'Planning Commission Approval' in content:_checkboxes(p,[is_pc,not is_pc,True,True,renewal])
    _clear_all_red_run_colors(doc)
    # Also include controls/textboxes omitted by python-docx's high-level runs.
    for color in doc.element.iter(qn('w:color')):
        shade=color.get(qn('w:val'),'').upper()
        if len(shade)==6 and shade.endswith('0000') and shade!='000000':color.set(qn('w:val'),'000000')
    if role=='routing':
        prepare_routing_form(doc)
        # The City variants have different attachment lists; address the
        # named checkboxes after the layout helper separates their rows.
        for p in doc.element.iter(qn('w:p')):
            content=''.join(t.text or '' for t in p.iter(qn('w:t')))
            if 'Original Ordinance' in content:_checkboxes(p,[wizard.endswith('_renewal')])
    out=Path(output);out.parent.mkdir(parents=True,exist_ok=True);doc.save(out);unlock_docx(out)
    return True
