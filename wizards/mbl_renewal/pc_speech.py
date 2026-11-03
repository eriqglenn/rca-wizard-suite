"""Renewal speech fields and reader; renewal flags never supply petition support."""
from wizards.mbl.pc_speech import speech_context, RECOMMENDATIONS
from wizards.mbl.pc_data import read_pc_facts
from wizards.mls import pc_speech as common

FIELDS=common.FIELDS+[
    ('ordinance','Previous ordinance number',('Original_Ordinance_No',)),
    ('adoption_date','Previous ordinance adoption date',('Original_Passage_Date','Original_Adoption_Date','Original_Council_Passage_Date')),
]

def defaults(facts,preferences):
    values=common.defaults(facts,preferences)
    for key,_,aliases in FIELDS[len(common.FIELDS):]:
        values[key]=next((str(facts['case_info'][k]) for k in aliases if facts['case_info'].get(k)),'')
        if not values[key]:values[key]=preferences.get(key,'')
    values.update(preferences.get('speech',{}))
    return values

def read_speech_facts(resolved,source=None):
    return read_pc_facts(resolved,source,'mbl_renewal')
