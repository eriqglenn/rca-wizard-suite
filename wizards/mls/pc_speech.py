"""Fill the City's MLS PC speech using saved calculation facts, not AI."""
from __future__ import annotations
from datetime import datetime
from decimal import Decimal
from pathlib import Path
import json
import re
from docx import Document
from docx.oxml.ns import qn
from core.paths import templates_dir
from core.boundary_map import fingerprint
from .pc_data import ELIGIBLE, number, percent, read_pc_facts

FIELDS = [
    ('agenda', 'Agenda item', ('PC_Agenda_Number',)),
    ('presenter', 'Presenter', ('PC_Presenter', 'Planner_Name')),
    ('neighborhood', 'Neighborhood', ('Neighborhood_Name', 'Neighborhood')),
    ('district', 'Council district', ('Council_District',)),
    ('super_neighborhood', 'Super Neighborhood', ('Super_Neighborhood', 'Super_Neighborhood_Name')),
    ('subdivision', 'Subdivision', ('Subdivision_Name',)),
    ('plat_year', 'Subdivision plat year', ('Subdivision_Plat_Year',)),
    ('earliest_year', 'Earliest house year', ('Earliest_House_Year', 'Earliest_Construction_Year')),
    ('protest_count', 'Timely written protests', ('Timely_Protest_Count', 'Protest_Count')),
]
RECOMMENDATIONS = {
    'Approval': 'forward the application to City Council for approval',
    'Disapproval': 'disapprove the application',
    'Deferral': 'defer the application',
}

def defaults(facts, preferences):
    """Reuse report details, but never confuse neighborhood history with its name."""
    ci=facts['case_info']; result={}
    for key, _label, keys in FIELDS:
        result[key]=next((str(ci[k]).strip() for k in keys if ci.get(k) not in (None,'')), '')
    for key in ('agenda','plat_year','protest_count'):
        if not result[key] and preferences.get(key) not in (None,''):result[key]=str(preferences[key])
    # PC-report notice transmission is not proof of mailed owner notifications.
    result.update(notice_mailed=False,recommendation='Approval')
    result.update(preferences.get('speech',{}))
    return result

def read_speech_facts(resolved, source=None):
    try:
        facts=read_pc_facts(resolved,source)
    except (TypeError, StopIteration) as error:
        raise ValueError('The MLS calculation sheet needs a calculated proposed minimum and property rows in Step 2. Save it in Excel and try again.') from error
    if facts['minimum']<=0 or facts['eligible_count']==0:
        raise ValueError('Step 2 needs a positive proposed minimum and at least one eligible lot for the PC speech.')
    return facts

def speech_context(facts, options):
    notes=list(facts['warnings'])
    def missing(label):
        notes.append('Complete or confirm '+label+'.')
        return '[Enter '+label+']'
    def field(key,label):
        value=str(options.get(key,'') or '').strip()
        if not value or (key!='agenda' and re.fullmatch(r'[Xx, ]+',value)) or value.startswith('#'):
            return missing(label)
        return value
    def year(key,label):
        value=field(key,label)
        if value.startswith('['):return value
        n=number(value)
        if n is None or n!=int(n) or not 1600<=n<=datetime.now().year:
            return missing(label+' (four-digit year)')
        return str(int(n))
    eligible=[r for r in facts['rows'] if r['land_use'] in ELIGIBLE]
    complete=all(r['land_use'] in ELIGIBLE|{'EXC'} and r['area'] is not None
                 and (r['area']>0 or r['land_use']=='EXC') for r in facts['rows'])
    excluded=facts['excluded_count']
    area_label='eligible application land area' if excluded else 'application land area'
    lot_label='eligible lots' if excluded else 'lots'
    counts={lu:sum(r['land_use']==lu for r in eligible) for lu in ELIGIBLE}
    total=sum(Decimal(str(r['area'])) for r in eligible) if complete else None
    if complete:
        sfr=counts['SFR'];com=counts['COM']
        com_area=sum(Decimal(str(r['area'])) for r in eligible if r['land_use']=='COM')
        land=(f'The application area consists of {sfr} single-family residential '+('lot' if sfr==1 else 'lots')+
              f', representing {percent(sfr,len(eligible))}% of the {lot_label} in the application area, and '
              f'{com} commercial '+('lot' if com==1 else 'lots')+f', representing {percent(com_area,total)}% of the {area_label}.')
        others=[]
        for lu,label in [('MF','multifamily residential'),('VAC','vacant')]:
            n=counts[lu]
            if n:others.append(f'{n} {label} '+('lot' if n==1 else 'lots'))
        if others:land+=' The area also includes '+' and '.join(others)+'.'
        qualifying=sum(Decimal(str(r['area']))>=Decimal(str(facts['minimum'])) for r in eligible)
    else:
        land='[Review Step 2 land uses and lot areas before presenting the land-use figures.]'
        qualifying='[Review Step 2]';notes.append('Correct incomplete or invalid Step 2 lot areas / land uses.')
    if excluded:
        land+=f' {excluded} excluded '+('lot is' if excluded==1 else 'lots are')+' outside these eligibility calculations.'
    if complete and all(r['support_known'] for r in eligible):
        signed=[r for r in eligible if r['signed']]
        signed_area=sum(Decimal(str(r['area'])) for r in signed)
        support=(f'The applicant provided evidence of support from owners of {len(signed)} of the {len(eligible)} '
                 f'{lot_label}, representing {percent(signed_area,total)}% of the '+
                 ('total eligible land area.' if excluded else 'total land area.'))
    else:
        support='[Review Step 2 support entries and lot areas before presenting support totals.]'
        notes.append('Confirm the support totals in Step 2.')
    protest=number(options.get('protest_count',''))
    if protest is None or protest<0 or protest!=int(protest):
        protest_text='[Confirm the number of timely written protests received.]'
        notes.append('Confirm the number of timely written protests received (enter 0 if none).')
    elif protest==0:protest_text='No timely written protests were received by staff.'
    elif protest==1:protest_text='One timely written protest from a property owner was received by staff.'
    else:protest_text=f'{int(protest)} timely written protests from property owners were received by staff.'
    notice='Notification was mailed to all property owners in the application area.'
    if options.get('notice_mailed') is not True:
        notice='[Confirm that notification was mailed to all property owners in the application area.]'
        notes.append('Confirm that owner notifications were mailed.')
    subdivision=field('subdivision','subdivision name')
    subphrase=(subdivision if re.search(r'\bsubdivision\b',subdivision,re.I) else subdivision+' Subdivision')
    recommendation=RECOMMENDATIONS.get(options.get('recommendation','Approval'))
    if recommendation is None:recommendation=missing('staff recommendation')
    boundary=facts['boundary'].rstrip(' .')
    if any(not str(v).strip() for v in [facts['street'],facts['block'],*facts['cross_streets']]) or not facts['blockfaces']:
        boundary='[Complete the boundary in Step 2]';notes.append('Complete the street, block range, sides and cross streets in Step 2.')
    ctx=dict(boundary=boundary,agenda=field('agenda','agenda item'),presenter=field('presenter','presenter name'),
             minimum=facts['minimum_text'],neighborhood=field('neighborhood','neighborhood name'),
             district=field('district','council district'),super_neighborhood=field('super_neighborhood','Super Neighborhood'),
             subdivision_phrase=subphrase,plat_year=year('plat_year','subdivision plat year'),
             earliest_year=year('earliest_year','earliest house year'),land_use_sentence=land,
             qualifying_count=str(qualifying),qualifying_noun='lot' if qualifying==1 else 'lots',
             support_sentence=support,notice_sentence=notice,protest_sentence=protest_text,recommendation=recommendation)
    return ctx,list(dict.fromkeys(notes))

def fill_speech(context):
    from core.templates import template_path
    wizard_key=context.get('wizard_key','mls')
    doc=Document(template_path(wizard_key,'pc_speech_template',{'mls':'MLS','mls_renewal':'MLS_REN','mbl':'MBL','mbl_renewal':'MBL_REN'}[wizard_key]+'_PC_Speech.docx'))
    for node in doc.part.element.iter(qn('w:t')):
        node.text=re.sub(r'\{\{(\w+)\}\}',lambda m:str(context[m[1]]),node.text or '')
    return doc

def speech_text(context):
    return '\n\n'.join(p.text for p in fill_speech(context).paragraphs if p.text)

def generate_speech(source,facts,options,output,export_pdf=True):
    if fingerprint(source)!=facts['workbook_hash']:
        raise ValueError('The workbook changed. Close PC Speech and reopen it to read the latest saved values.')
    reader=speech_context
    if facts.get('wizard_key')=='mls_renewal':
        from wizards.mls_renewal.pc_speech import speech_context as reader
    if facts.get('wizard_key') in ('mbl','mbl_renewal'):
        from wizards.mbl.pc_speech import speech_context as reader
    context,notes=reader(facts,options)
    app=re.sub(r'[^A-Za-z0-9_-]+','_',facts['application']).strip('_') or 'Case'
    prefix={'mls':'MLS','mls_renewal':'MLS_REN','mbl':'MBL','mbl_renewal':'MBL_REN'}[facts.get('wizard_key','mls')]
    folder=Path(output)/f'{prefix}_{app}_PC_Speech_{datetime.now():%Y%m%d_%H%M%S_%f}'
    folder.mkdir(parents=True,exist_ok=False)
    word=folder/('PC_Speech_DRAFT.docx' if notes else 'PC_Speech.docx')
    fill_speech(context).save(word)
    pdf=None
    if export_pdf:
        from core.office_convert import office_to_pdf
        try:pdf=str(office_to_pdf(word,folder))
        except Exception as error:notes.append('Word speech saved. PDF export was unavailable: '+str(error))
    audit={'source_workbook':str(source),'workbook_hash':facts['workbook_hash'],
           'application':facts['application'],'options':options,'context':context,'review_notes':notes,
           'metric_basis':'Step 2 legal lot rows; EXC excluded; SFR and support by land area; building lines in feet and qualifying counts by developed lot. Renewal flags are not support.' if facts.get('wizard_key','').startswith('mbl') else 'Step 2 legal lot rows; EXC excluded; SFR by lot count, commercial by land area. Renewal support is not required.' if facts.get('wizard_key')=='mls_renewal' else 'Step 2 legal lot rows; EXC excluded; SFR by lot count, commercial and support by land area.'}
    (folder/'PC_Speech_Review.json').write_text(json.dumps(audit,indent=2),encoding='utf8')
    return {'docx':str(word),'pdf':pdf,'folder':str(folder),'notes':notes}
