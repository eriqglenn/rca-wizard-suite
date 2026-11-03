"""Building-line speech from the shared City speech, using building-line facts."""
from wizards.mls import pc_speech as common
from .pc_data import read_pc_facts
from .pc_report import support_sentence
from wizards.mls.pc_data import date_text

FIELDS=common.FIELDS
RECOMMENDATIONS=common.RECOMMENDATIONS
defaults=common.defaults


def read_speech_facts(resolved,source=None):
    return read_pc_facts(resolved,source,'mbl')


def speech_context(facts,options):
    ctx,notes=common.speech_context(facts,options)
    ctx['wizard_key']=facts['wizard_key'];ctx['application']=facts['application']
    ctx['land_use_sentence']=(f"The application area includes {facts['sfr_count']} single-family residential lots, representing "
        f"{facts['sfr_area_percent']}% of the eligible application land area, and {facts['vacant_count']} vacant "+('lot.' if facts['vacant_count']==1 else 'lots.'))
    for lu,label in [('COM','commercial'),('MF','multifamily residential')]:
        n=sum(r['land_use']==lu for r in facts['rows'])
        if n:ctx['land_use_sentence']+=f" There {'is' if n==1 else 'are'} {n} {label} "+('lot.' if n==1 else 'lots.')
    if facts['excluded_count']:ctx['land_use_sentence']+=f" {facts['excluded_count']} EXC "+('lot is' if facts['excluded_count']==1 else 'lots are')+' outside the eligibility calculations.'
    ctx['minimum_sentence']=(f"A building line of at least {facts['minimum_text']} feet is recorded on "
        f"{facts['qualifying_count'] if facts['qualifying_count'] is not None else '[review]'} of {facts['developed_count']} developed lots "
        f"({facts['developed_percent']}%). Vacant lots are excluded from that denominator.")
    ctx['support_sentence']=support_sentence(facts)
    ctx['renewal_history']=''
    if facts['wizard_key']=='mbl_renewal':
        ci=facts['case_info']
        ordinance=str(options.get('ordinance',ci.get('Original_Ordinance_No',''))).strip()
        adoption=date_text(options.get('adoption_date',ci.get('Original_Passage_Date','')))
        if not ordinance:ordinance='[Enter previous ordinance number]';notes.append('Complete the previous ordinance number.')
        if not adoption:adoption='[Enter previous ordinance adoption date]';notes.append('Complete the previous ordinance adoption date.')
        ctx['renewal_history']=f'The previous ordinance, {ordinance}, was adopted on {adoption}.'
    notes.append('Confirm that developed-lot rows represent the structures being described; the worksheet does not independently count buildings.')
    return ctx,list(dict.fromkeys(notes))
