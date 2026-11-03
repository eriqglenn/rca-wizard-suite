"""Renewal presentation using the common City speech format and renewal facts."""
from wizards.mls.pc_speech import FIELDS as COMMON_FIELDS, RECOMMENDATIONS, defaults as common_defaults
from wizards.mls.pc_speech import speech_context as common_context, speech_text, generate_speech
from .pc_data import read_pc_facts

FIELDS = COMMON_FIELDS + [
    ('ordinance', 'Previous ordinance number', ('Original_Ordinance_No',)),
    ('adoption_date', 'Previous ordinance adoption date', ('Original_Passage_Date', 'Original_Adoption_Date', 'Original_Council_Passage_Date')),
]

def defaults(facts, preferences):
    values = common_defaults(facts, preferences)
    for key, _, names in FIELDS[len(COMMON_FIELDS):]:
        values[key] = next((facts['case_info'][n] for n in names if facts['case_info'].get(n)), '')
        values[key] = preferences.get('speech', {}).get(key, values[key])
    return values

def read_speech_facts(path, source=None):
    return read_pc_facts(path, source)

def speech_context(facts, options):
    ctx, notes = common_context(facts, options)
    notes = [n for n in notes if n != 'Confirm the support totals in Step 2.']
    ctx['support_sentence'] = 'Evidence of petition support is not required for this renewal.'
    ctx['wizard_key'] = 'mls_renewal'
    ctx['application'] = facts['application']
    for key, label in [('ordinance', 'previous ordinance number'), ('adoption_date', 'previous ordinance adoption date')]:
        ctx[key] = str(options.get(key, '') or '').strip()
        if not ctx[key]:
            ctx[key] = '[Enter ' + label + ']'; notes.append('Complete ' + label + '.')
    return ctx, notes
