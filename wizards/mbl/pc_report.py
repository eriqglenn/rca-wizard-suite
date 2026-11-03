"""Fill the SharePoint building-line report with the matching Step 2 measures."""
from wizards.mls.pc_report import report_context as common_context, words_count
from wizards.mls.pc_data import date_text


def report_context(facts, options):
    ctx,notes=common_context(facts,options)
    renewal=facts['wizard_key']=='mbl_renewal';ci=facts['case_info']
    ctx['wizard_key']=facts['wizard_key']
    ctx['sfr_sentence']=(f"The application includes {words_count(facts['sfr_count'])} single-family residential lots, representing "
        f"{facts['sfr_area_percent']}% of the eligible application land area. {words_count(facts['vacant_count'],True)} "+('lot is' if facts['vacant_count']==1 else 'lots are')+' vacant.')
    for lu,label in [('COM','commercial'),('MF','multifamily residential')]:
        n=sum(r['land_use']==lu for r in facts['rows'])
        if n:ctx['sfr_sentence']+=f" There {'is' if n==1 else 'are'} {words_count(n)} {label} "+('lot.' if n==1 else 'lots.')
    if facts['excluded_count']:ctx['sfr_sentence']+=f" {words_count(facts['excluded_count'],True)} EXC "+('lot is' if facts['excluded_count']==1 else 'lots are')+' outside the eligibility calculations.'
    ctx['support_sentence']=support_sentence(facts)+' '+ctx['protest_sentence']
    if facts['signed_count'] is None:
        notes.append('No complete set of explicit petition-support entries is available. Renewal compliance flags are not support signatures.')
    ctx['minimum_sentence']=(f"A building line of at least {facts['minimum_text']} feet is recorded on {words_count(facts['qualifying_count'])} "
        f"of {words_count(facts['eligible_count'])} eligible lots. "
        f"{words_count(facts['below_count'],True)} developed lots have a building line below that minimum. "
        f"The application also includes {words_count(facts['vacant_count'])} vacant "+('lot.' if facts['vacant_count']==1 else 'lots.'))
    ctx['structure_sentence']=(f"A building line of at least {facts['minimum_text']} feet is recorded on {words_count(facts['qualifying_count'])} "
        f"of {words_count(facts['developed_count'])} developed lots ({facts['developed_percent']}%). Vacant and EXC lots are not included in this denominator.")
    history=str(options.get('preservation',ci.get('Preservation_Finding',''))).strip()
    # Replace only the shared MLS default; planner-entered findings are retained.
    if not history:
        old=f"The establishment of a {facts['minimum_text']} sf minimum lot size will preserve the lot size character of the area."
        ctx['history_sentence']=ctx['history_sentence'].replace(old,f"[Confirm how the {'renewal' if renewal else 'establishment'} of a {facts['minimum_text']}-foot minimum building line will preserve the building-line character of the area.]")
        notes.append('Confirm the preservation finding; it is a planner judgment.')
    ctx['renewal_history']=''
    if renewal:
        ordinance=str(options.get('ordinance',ci.get('Original_Ordinance_No',''))).strip()
        adoption=date_text(options.get('adoption_date',ci.get('Original_Passage_Date','')))
        if not ordinance:ordinance='[Enter previous ordinance number]';notes.append('Complete the previous ordinance number.')
        if not adoption:adoption='[Enter previous ordinance adoption date]';notes.append('Complete the previous ordinance adoption date; the effective date is separate.')
        ctx['renewal_history']=f'The previous ordinance, {ordinance}, was adopted on {adoption}.'
        ctx['criteria_note']='[Review the source template’s renewal criteria against the renewal calculation sheet before presenting this report.]'
    else:ctx['criteria_note']=''
    notes.append('Developed-lot counts come from Step 2 rows, not an independent structure count. Confirm this denominator if a property contains multiple structures.')
    return ctx,list(dict.fromkeys(notes))


def support_sentence(facts):
    if facts['signed_count'] is None:
        return 'Petition support is not recorded in the calculation sheet. [Confirm the applicable support finding.]'
    return (f"Owners signed in support for {words_count(facts['signed_count'])} of {words_count(facts['eligible_count'])} eligible lots, "
            f"representing {facts['support_area_percent']}% of the eligible application land area.")
