"""City MLS Renewal PC report fields, independent of new-application support rules."""
from wizards.mls.pc_data import date_text, number
from wizards.mls.pc_report import words_count


def report_context(facts, options):
    ci = facts['case_info']; notes = list(facts['warnings'])
    def field(key, ci_key=None):
        value = str(options.get(key, ci.get(ci_key or key, '')) or '').strip()
        if not value:
            notes.append('Complete ' + key.replace('_', ' ') + '.')
            return '[Enter ' + key.replace('_', ' ') + ']'
        return value
    protest = number(options.get('protest_count', ci.get('Protest_Count')))
    if protest is None or protest < 0 or protest != int(protest):
        protest_sentence = '[Confirm timely written protests.]'; notes.append('Confirm the timely written protest count.')
    elif protest == 0: protest_sentence = 'No timely written protests were filed.'
    else: protest_sentence = words_count(int(protest), True) + (' timely written protest was filed.' if protest == 1 else ' timely written protests were filed.')
    notice = 'All property owners were provided notice stating that a protest must be filed within thirty (30) days.'
    if options.get('notice_sent') is not True:
        notice = '[Confirm owner notice and the thirty-day protest period.]'; notes.append('Confirm owner notice.')
    hearing_notice = 'Public notice of the public hearing was transmitted to all property owners on the blockfaces.'
    if options.get('hearing_notice_sent') is not True:
        hearing_notice = '[Confirm that public-hearing notice was transmitted to all property owners.]'; notes.append('Confirm public-hearing notice separately from the owner renewal notice.')
    variance = facts.get('variance')
    if variance:
        variance_sentence = f"Excluding corner lots, {words_count(variance['within'])} of {words_count(variance['count'])} lots ({variance['percent']}%) vary by no more than 10 percent from the average lot size of {variance['average']} square feet."
    else: variance_sentence = field('variance_statement', 'PC_Lot_Size_Variance_Statement')
    area = facts['boundary']; minimum = facts['minimum_text']
    ctx = dict(agenda=field('agenda','PC_Agenda_Number'), application=facts['application'], boundary=area,
               hearing_date=date_text(field('hearing_date','PC_Hearing_Date')), minimum=minimum,
               ordinance=field('ordinance','Original_Ordinance_No'),
               adoption_date=date_text(field('adoption_date','Original_Passage_Date')),
               condition_statement=field('condition_statement','Renewal_Block_Condition_Statement'),
               protest_sentence=protest_sentence, all_lots=words_count(facts['all_count']),
               blockface_sentence=f"The application comprises {words_count(facts['blockfaces'])} blockface(s), the {facts['sides']} of {facts['street']}.",
               sfr_sentence=f"{words_count(facts['sfr_count'], True)} of {words_count(facts['eligible_count'])} eligible lots are single-family residential ({facts['sfr_lot_percent']}% by lot count and {facts['sfr_area_percent']}% by land area)." + (f" {facts['excluded_count']} EXC lot(s) are excluded from these figures." if facts['excluded_count'] else ''),
               support_sentence='Evidence of support is not required per the renewal report template. ' + notice + ' ' + protest_sentence,
               minimum_sentence=f"A lot size of at least {minimum} square feet exists on {words_count(facts['qualifying_count'])} eligible lots in the application area.",
               variance_sentence=variance_sentence, notice_sentence=hearing_notice)
    notes.append('The SharePoint source uses SFR land area in Procedures and SFR lot count in Staff Analysis. Both figures are reported separately; confirm the applicable criterion before finalizing.')
    return ctx, list(dict.fromkeys(notes))
