"""MBL Planner Wizard (v9.0, 2026 City templates).

Fills the official City of Houston Special Minimum Building Line Block
(new application) templates from the MBL calculation workbook (Step 1 GIS
data, Step 2 calculation sheet, Case Info tab) and assembles the RCA packet.

Documents generated (templates in templates/mbl, filled through the shared
wording-based engine in core/docx_merge.py):

    Owner notification letters       mail merge, map after each letter
    Notice of public hearing         support under 51%: PC hearing notice to all owners
    Protest letter to owners         public-hearing notice when a protest is filed
    Notice of PC rescheduling        new hearing date to all owners
    Applicant sign letter            completion / sign-posting letter
    Director approval                2026 criteria form (support 51% or more, no protest)
    Planning Commission approval     criteria form with findings (PC path)
    Routing form                     2026 coversheet routing form (content controls)
    OneMeeting staff report          Director path, support-under-51% path or protest path
    Draft ordinance                  City ordinance template (token engine, see mbl_mailout_engine)
    RCA packet PDF                   routing, approval, map, Step 2 staff analysis, plat, tax map, HCAD
"""
import tkinter as tk
import datetime
import math
import re
from tkinter import filedialog, messagebox
from pathlib import Path
import tempfile, os, shutil, subprocess, sys, copy, threading, time, urllib.request
from copy import deepcopy
from openpyxl import load_workbook
from core.routing_form import prepare_routing_form, routing_pdf_one_page
from docx import Document
from docx.oxml.ns import qn

from core.docx_unlock import unlock_docx
from core.embedding import EmbeddedWizardMixin
from core.scrollable import ScrollableContent
from core.workbooks import calculation_template_path, WorkbookReadCache
from core import session as core_session
from core.casefolder import maybe_case_folder
from core.errorlog import log_exception
from core.preflight import run_preflight, build_preflight_panel, build_snapshot_panel
from core.paths import templates_dir, sample_input_dir, HCAD_DOWNLOADER
from core.templates import template_path
from core import launch as core_launch
from core.pdf_condense import condense_hcad_pdfs
from core.docx_merge import (
    _RunCursor, _append_letter_as_own_section, _cap, _clean, _clear_all_red_run_colors,
    _compact_notification_first_page, _fill_phone, _fix_mailto_links, _fmt_pct, _fmt_pct_num,
    _insert_map, _iter_all_paragraphs, _map_to_png, _norm, _num_words, _parse_date,
    _replace_in_runs, _replace_runs_contextual, _replace_text_nodes, _restart_list_numbering,
    _set_exact_para_preserve_first_run, _set_paragraph_checkboxes, _set_run_text_preserve_xml,
    _blacken_run, _to_int, _has_hyperlink, _EMAIL_PLACEHOLDERS, _EXCEL_ERRORS,
)
# Shared pieces of the other 2026 wizards: routing-form content controls (MLS),
# Step 1 reader and PDF helpers (MLS Renewal), criteria-line filler (MBL Renewal).
from wizards.mls.templates_2026 import _fill_sdts_in_order
from wizards.mls_renewal.wizard import _step1_reader, _merge_pdfs, _ensure_pdf, _convert_office_to_pdf
from wizards.mbl_renewal.wizard import _fill_runs, _replace_line_in_paragraph
# The draft ordinance still runs on the token engine (the City's 2026 ordinance
# file is password-protected); the Step 2 staff analysis and HCAD export are
# unchanged from v8.
from .mbl_mailout_engine import (
    generate_draft_ordinance as _legacy_generate_draft_ordinance,
    export_mbl_step2_staff_analysis_pdf, export_hcad_address_list,
)

WIZARD_VERSION = 'v9.0'
STEP2_SHEET = 'Step 2 - Main Calculation Sheet'
STEP1_SHEET = 'Step 1 - Input Properties Info'
_WIZARD_KEY = 'mbl'


# =============================================================================
# Workbook readers
# =============================================================================
def _to_float(v, default=0.0):
    try:
        if v is None:
            return default
        if isinstance(v, str):
            s = v.strip().replace(',', '').replace('%', '')
            if not s:
                return default
            return float(s)
        return float(v)
    except Exception:
        return default


def _bl_text(v):
    """Building line in feet: 15 -> '15', 12.5 -> '12.5'."""
    f = _to_float(v)
    if f <= 0:
        return ''
    return str(int(round(f))) if abs(f - round(f)) < 0.01 else f"{f:g}"


def _get_step2(wb):
    """Read the City's MBL calculation sheet: case header (rows 2-6), the
    support / SFR summary (rows 8-15) and the property table (row 17 on)."""
    ws = wb[STEP2_SHEET] if STEP2_SHEET in wb.sheetnames else wb[wb.sheetnames[1]]
    rows = []
    for r in range(17, min(ws.max_row, 140) + 1):
        desc = _clean(ws.cell(r, 1).value)
        lot = _to_float(ws.cell(r, 4).value)
        bl = _to_float(ws.cell(r, 5).value)
        if not desc and lot <= 0 and bl <= 0:
            continue
        rows.append({'desc': desc, 'land_use': _clean(ws.cell(r, 2).value).upper(),
                     'signed': _clean(ws.cell(r, 3).value).upper().startswith('Y'),
                     'lot': lot, 'bl': bl, 'row': r})
    if not any(x['desc'] for x in rows):
        # Stale formula cache: the Step 2 mirror formulas are empty.  Rebuild the
        # table from the typed Step 1 GIS columns plus the typed Step 2 columns
        # (C = signed, E = constructed building line).
        try:
            s1, gv = _step1_reader(wb)
            rebuilt = []
            for i in range(0, 125):
                r1, r2 = 2 + i, 17 + i
                desc = gv(r1, 'Property Description')
                legal = gv(r1, 'Legal_Dscr_1')
                num = gv(r1, 'Str_num')
                if not desc and not legal and not num:
                    continue
                lu = gv(r1, 'Land Use').upper()
                if not lu:
                    grp = gv(r1, 'group_dscr').lower()
                    lu = 'SFR' if 'single-family' in grp else ('VAC' if 'undeveloped' in grp else (
                        'COM' if 'commercial' in grp else ('MF' if 'multi-family' in grp else 'EXC')))
                rebuilt.append({'desc': desc or f"{num} ({legal})", 'land_use': lu,
                                'signed': _clean(ws.cell(r2, 3).value).upper().startswith('Y'),
                                'lot': _to_float(gv(r1, 'Lot Size') or gv(r1, 'Total_Land_Area')),
                                'bl': _to_float(ws.cell(r2, 5).value), 'row': r2})
            if rebuilt:
                rows = rebuilt
        except Exception:
            pass
    developed = [x['bl'] for x in rows if x['bl'] > 0]
    min_bl = _to_float(ws['L16'].value)
    if min_bl <= 0:
        m = re.match(r'\s*([\d.,]+)', _clean(ws['E4'].value))
        min_bl = _to_float(m.group(1)) if m else 0
    if min_bl <= 0 and developed:
        # the sheet's rule: the largest line that at least 70% of the developed lots meet
        srt = sorted(developed, reverse=True)
        min_bl = srt[max(0, math.ceil(0.7 * len(srt)) - 1)]
    total = _to_int(ws['A15'].value) or len(rows)
    sfr_count = _to_int(ws['A11'].value) or sum(1 for x in rows if x['land_use'] == 'SFR')
    total_area = _to_float(ws['B8'].value) or sum(x['lot'] for x in rows)
    signed_area = _to_float(ws['D8'].value) or sum(x['lot'] for x in rows if x['signed'])
    support = _to_float(ws['F8'].value)
    if support <= 0 and total_area:
        support = signed_area / total_area
    sfr_pct = _to_float(ws['H11'].value)
    if sfr_pct <= 0 and total:
        sfr_pct = sfr_count / total
    developed_n = _to_int(ws['AC17'].value) or len(developed)
    qual = sum(1 for b in developed if min_bl and b >= min_bl)
    case = {
        'Application_Number': _clean(ws['B2'].value),
        'Date_Received': _clean(ws['B3'].value),
        'Date_Complete': _clean(ws['D3'].value),
        'Street_Name': _clean(ws['B4'].value),
        'Block_Description': _clean(ws['D4'].value),
        'Cross_Street_1': _clean(ws['B5'].value),
        'Cross_Street_2': _clean(ws['D5'].value),
        'Side_1': _clean(ws['B6'].value),
        'Side_2': _clean(ws['D6'].value),
        'Minimum_Building_Line': min_bl,
        'Minimum_Building_Line_Text': _bl_text(min_bl),
        'Support_Percentage': _fmt_pct(support) if support else '',
        'SFR_Percentage': _fmt_pct(sfr_pct) if sfr_pct else '',
        'SFR_Count': sfr_count,
        'Total_Lot_Count': total,
        'Developed_Lot_Count': developed_n,
        'Qualifying_Lot_Count': qual,
        'Building_Line_Percentage': _fmt_pct((qual / developed_n) if developed_n else 0),
        'Total_Area': total_area,
        'Signed_Area': signed_area,
        'Qualifies': _clean(ws['E3'].value),
        'Property_Rows': rows,
    }
    if case['Application_Number'] in ('0', '0.0'):
        case['Application_Number'] = ''
    for k in ('Date_Received', 'Date_Complete'):
        if case[k] in ('MM/DD/YYYY',):
            case[k] = ''
    sides = [x.strip().lower() for x in (case['Side_1'], case['Side_2']) if x.strip()]
    case['Side_Text'] = ' and '.join(sides)
    case['Side_Label'] = (case['Side_Text'] + (' sides' if len(sides) > 1 else ' side')) if sides else ''
    block = case['Block_Description']
    case['Area_Line_1'] = f"{block}, {case['Side_Label']}," if case['Side_Label'] else (f"{block}," if block else '')
    c1, c2 = case['Cross_Street_1'], case['Cross_Street_2']
    case['Area_Line_2'] = f"between {c1} and {c2}" if (c1 or c2) else ''
    return case


def _get_case_info(wb):
    d = {}
    if 'Case Info' in wb.sheetnames:
        ws = wb['Case Info']
        for r in range(1, ws.max_row + 1):
            k = ws.cell(r, 1).value
            if k:
                d[str(k).strip()] = ws.cell(r, 2).value
    try:
        s2 = _get_step2(wb)
        overlay = {
            'Application_Number': s2.get('Application_Number', ''),
            'Application_Type': 'Special Minimum Building Line Block Application',
            'Date_Received': s2.get('Date_Received', ''),
            'Date_Complete': s2.get('Date_Complete', ''),
            'Street_Name': s2.get('Street_Name', ''),
            'Block_Description': s2.get('Block_Description', ''),
            'Side_Of_Street': s2.get('Side_Text', ''),
            'Cross_Streets': (s2.get('Area_Line_2', '') or '').replace('between ', '').strip(),
            'Minimum_Building_Line': s2.get('Minimum_Building_Line_Text', ''),
            'Support_Percentage': s2.get('Support_Percentage', ''),
            'SFR_Percentage': s2.get('SFR_Percentage', ''),
            'SFR_Count': s2.get('SFR_Count', ''),
            'Total_Lot_Count': s2.get('Total_Lot_Count', ''),
            'Developed_Lot_Count': s2.get('Developed_Lot_Count', ''),
            'Qualifying_Lot_Count': s2.get('Qualifying_Lot_Count', ''),
            'Qualifies': s2.get('Qualifies', ''),
        }
        for k, v in overlay.items():
            sv = '' if v is None else str(v).strip()
            if sv and sv.upper() not in _EXCEL_ERRORS:
                d[k] = v
    except Exception:
        pass
    return d


def _get_owners(wb):
    """One mail-out record per Step 1 property, including shared recipients."""
    ws, gv = _step1_reader(wb)
    owners = []
    for r in range(2, ws.max_row + 1):
        owner = gv(r, 'OWNER_MAILTO') or gv(r, 'OWNER_LIST')
        addr1 = gv(r, 'Mail_Addr_1')
        if not owner or not addr1 or owner in ('0', '<Null>'):
            continue
        # A shared owner or mailing address must not suppress another property.
        owners.append({'Excel_Row':str(r),
            'Owner_Mail': owner, 'Mail_Addr_1': addr1, 'Mail_Addr_2': gv(r, 'Mail_Addr_2'),
            'Mail_City': gv(r, 'Mail_City'), 'Mail_State': gv(r, 'Mail_State'), 'Mail_Zip': gv(r, 'Mail_Zip'),
            'Property_Description': gv(r, 'Property Description') or gv(r, 'Site_addr_1'),
        })
    from core.lot_divisions import collapse_lot_notifications
    return collapse_lot_notifications(wb,owners)


def _unwrap_smart_tags(doc):
    """Word address/date smart tags (w:smartTag) wrap runs that python-docx
    does not list as paragraph runs; the City's converted .doc forms carry them
    around sample addresses.  Unwrap them so every run can be filled."""
    body = doc.element.body
    changed = True
    while changed:
        changed = False
        for st in list(body.iter(qn('w:smartTag'))):
            parent = st.getparent()
            if parent is None:
                continue
            idx = parent.index(st)
            for child in list(st):
                if child.tag == qn('w:smartTagPr'):
                    continue
                parent.insert(idx, child)
                idx += 1
            parent.remove(st)
            changed = True


def _protest_sentence(n, zero_text):
    if n <= 0:
        return zero_text
    if n == 1:
        return 'One (1) protest was received.'
    return f'{_count_phrase(n)} protests were received.'


def _get_context(workbook_path):
    wb = load_workbook(workbook_path, data_only=True)
    ci = _get_case_info(wb)
    s = _get_step2(wb)
    app = str(s.get('Application_Number') or _clean(ci.get('Application_Number')) or '').strip()
    if app in ('0', '0.0'):
        app = ''
    block = s.get('Block_Description', '')
    side_label = s.get('Side_Label', '')
    c1, c2 = s.get('Cross_Street_1', ''), s.get('Cross_Street_2', '')
    between = f"between {c1} and {c2}" if (c1 or c2) else ''
    area = ', '.join(x for x in [block, side_label, between] if x)
    caption = (f"{side_label} of the {block}" if side_label else block)
    if between:
        caption += f", {between}"
    min_bl = s.get('Minimum_Building_Line') or 0
    min_text = s.get('Minimum_Building_Line_Text') or ''
    total = _to_int(s.get('Total_Lot_Count'))
    sfr = _to_int(s.get('SFR_Count'))
    dev = _to_int(s.get('Developed_Lot_Count'))
    qual = _to_int(s.get('Qualifying_Lot_Count'))
    support_num = _fmt_pct_num(s.get('Support_Percentage')) if s.get('Support_Percentage') else ''
    sfr_num = _fmt_pct_num(s.get('SFR_Percentage')) if s.get('SFR_Percentage') else ''
    protest_n = _to_int(ci.get('Protest_Count'), 0)
    hearing = _clean(ci.get('PC_Hearing_Date'))
    hearing_date = _parse_date(hearing)
    plat_year = _clean(ci.get('Subdivision_Plat_Year'))
    try:
        support_val = float(support_num) if support_num else None
    except Exception:
        support_val = None
    if protest_n > 0:
        variant = 'protest'
    elif support_val is not None and support_val < 51:
        variant = 'less51'
    else:
        variant = 'director'
    context = {**{k: _clean(v) for k, v in ci.items()}, **s}
    context.update({
        'Application_Number': app or 'XXX',
        'Area_Description': area,
        'Area_Description_Upper': area.upper(),
        'Ordinance_Caption': caption,
        'Between_Streets': between,
        'Applicant_Name': _clean(ci.get('Applicant_Name')) or 'Applicant Name',
        'Applicant_Address_Block': '\n'.join(x for x in [_clean(ci.get('Applicant_Address_1')), _clean(ci.get('Applicant_Address_2'))] if x) or 'Address',
        'Applicant_City': _clean(ci.get('Applicant_City')) or 'Houston',
        'Applicant_State': _clean(ci.get('Applicant_State')) or 'TX',
        'Applicant_Zip': _clean(ci.get('Applicant_Zip')) or '',
        'Minimum_Building_Line_Number': min_text or 'XX',
        'Minimum_Building_Line_Text': min_text or 'XX',
        'Minimum_Building_Line_Words': _num_words(int(round(min_bl))) if min_bl else 'X',
        'Support_Percentage_Number': support_num or 'XX',
        'SFR_Percentage_Number': sfr_num or 'XX',
        'Total_Lot_Count': str(total) if total else 'XX',
        'Total_Lot_Count_Words': _num_words(total) if total else 'XX',
        'SFR_Lot_Count': str(sfr) if sfr else 'X',
        'Developed_Lot_Count': str(dev) if dev else 'X',
        'Developed_Lot_Count_Words': _num_words(dev) if dev else 'X',
        'Qualifying_Lot_Count': str(qual) if qual else 'X',
        'Qualifying_Lot_Count_Words': _num_words(qual) if qual else 'X',
        'Building_Line_Percentage_Number': _fmt_pct_num(s.get('Building_Line_Percentage')) if dev else 'XX',
        'Protest_Count': protest_n,
        'Staff_Report_Variant': variant,
        'Council_Deadline': _clean(ci.get('Council_Deadline')) or '',
        'Council_District': _clean(ci.get('Council_District')) or '',
        'Council_Member': _clean(ci.get('Council_Member')).upper() or 'COUNCIL MEMBER NAME',
        'Approval_Path': _clean(ci.get('Approval_Path')) or 'Director Approval',
        'Planner_Name': _clean(ci.get('Planner_Name')) or 'Planner Name',
        'Planner_Position': _clean(ci.get('Planner_Position')) or 'Planner',
        'Planner_Phone': _clean(ci.get('Planner_Phone')) or '832-393-XXXX',
        'Planner_Email': _clean(ci.get('Planner_Email')) or 'Planner.Name@houstontx.gov',
        'Letter_Date': _clean(ci.get('Letter_Date')) or 'Date',
        'Protest_Letter_Date': _clean(ci.get('Protest_Letter_Date')) or _clean(ci.get('Letter_Date')) or 'Date',
        'Hearing_Letter_Date': _clean(ci.get('Hearing_Letter_Date')) or _clean(ci.get('Letter_Date')) or 'Date',
        'Reschedule_Letter_Date': _clean(ci.get('Reschedule_Letter_Date')) or _clean(ci.get('Letter_Date')) or 'Date',
        'Protest_Deadline': _clean(ci.get('Protest_Deadline')) or 'Month XX, XXXX',
        'PC_Hearing_Date': hearing or '[Date]',
        'PC_Hearing_Weekday': (hearing_date.strftime('%A') + ',') if hearing_date else 'Thursday,',
        'Director_Name': _clean(ci.get('Director_Name')) or 'Vonn Tran',
        'Applicant_Property_Address': _clean(ci.get('Applicant_Property_Address')) or 'XXXX XXX Street',
        'Subdivision_Name': _clean(ci.get('Subdivision_Name')) or 'SUBDIVISION NAME',
        'Sign_Count_Number': _clean(ci.get('Sign_Count_Number')) or '2',
        'Sign_Count_Words': _clean(ci.get('Sign_Count_Words')) or 'two',
        'Sign_Placement': _clean(ci.get('Sign_Placement')) or 'one sign at each corner',
        'Neighborhood_Age_Statement': _clean(ci.get('Neighborhood_Age_Statement')) or f"The subdivision was platted in {plat_year or 'XXXX'}.",
        'Structure_Age_Statement': _clean(ci.get('Structure_Age_Statement')) or 'The houses mostly originate from the 19XXs.',
        'Preservation_Finding': _clean(ci.get('Preservation_Finding')) or f"The establishment of a {min_text or 'XX'} foot minimum building line will preserve the building line character of the area.",
    })
    context['Qualifying_Phrase'] = (f"{_cap(context['Qualifying_Lot_Count_Words'])} ({context['Qualifying_Lot_Count']}) out of "
                                    f"{context['Developed_Lot_Count_Words']} ({context['Developed_Lot_Count']})")
    context['Qualifying_Sentence'] = (f"{context['Qualifying_Phrase']} developed properties (representing "
                                      f"{context['Building_Line_Percentage_Number']}% of the application area) have a building line of at least "
                                      f"{context['Minimum_Building_Line_Text']} feet.")
    context['PC_Narrative'] = ' '.join(x for x in [context['Neighborhood_Age_Statement'], context['Structure_Age_Statement'], context['Preservation_Finding']] if x)
    return context


# =============================================================================
# Owner letters (notification / public hearing / protest / PC rescheduling)
# =============================================================================
def _fill_owner_letter(doc, ctx, owner):
    _unwrap_smart_tags(doc)
    full = _norm(' '.join(p.text for p in doc.paragraphs))
    if 'received a protest' in full:
        letter_date = ctx['Protest_Letter_Date']
    elif 'rescheduled' in full:
        letter_date = ctx['Reschedule_Letter_Date']
    elif 'less than 51%' in full or 'public hearing will be held' in full:
        letter_date = ctx['Hearing_Letter_Date']
    else:
        letter_date = ctx['Letter_Date']
    min_bl = ctx['Minimum_Building_Line_Text']
    support = ctx['Support_Percentage_Number']
    name, phone, email = ctx['Planner_Name'], ctx['Planner_Phone'], ctx['Planner_Email']
    addr = owner.get('Mail_Addr_1', '')
    if owner.get('Mail_Addr_2') and owner.get('Mail_Addr_2') not in ('<Null>', '0'):
        addr += '\n' + owner['Mail_Addr_2']
    common = {
        '«Owner_Mail»': owner.get('Owner_Mail', ''),
        '«Mail_Addr_»': addr,
        '«Mail_City»': owner.get('Mail_City', ''),
        '«Mail_State»': owner.get('Mail_State', ''),
        '«Mail_Zip»': owner.get('Mail_Zip', ''),
        'XXX block of XX Street, XX and XX sides,': ctx['Area_Line_1'] or 'XXX block of XX Street, XX and XX sides,',
        'between XX and XX': ctx['Area_Line_2'] or 'between XX and XX',
    }
    for p in doc.paragraphs:
        t = _norm(p.text)
        ts = t.strip()
        if ts == 'DATE':
            _set_exact_para_preserve_first_run(p, letter_date)
            continue
        if '2:30' in t and 'DATE' in t:
            _fill_runs(p, exact_map={'DATE': ctx['PC_Hearing_Date']})
        elif 'protest' in t and ('DATE' in t or 'Month XX' in t):
            _fill_runs(p, exact_map={'DATE': ctx['Protest_Deadline'], 'Month XX, 2026': ctx['Protest_Deadline']},
                       contains_map={'Month XX, 2026': ctx['Protest_Deadline'], 'DATE': ctx['Protest_Deadline']})
        if 'public hearing will be held' in t:
            _fill_runs(p, exact_map={'X%': f"{support}%", 'X': min_bl, '[Date]': ctx['PC_Hearing_Date']},
                       contains_map={'[Date]': ctx['PC_Hearing_Date']})
        elif 'public hearing is scheduled' in t:
            _fill_runs(p, exact_map={'Thursday,': ctx['PC_Hearing_Weekday']},
                       contains_map={'[Date]': ctx['PC_Hearing_Date']})
        elif '% of property owners' in t:
            _fill_runs(p, exact_map={'XX': support})
        elif 'feet' in t:
            _fill_runs(p, exact_map={'XX': min_bl, 'X': min_bl, 'X feet': f"{min_bl} feet", 'XX feet': f"{min_bl} feet"})
        if 'Planner' in t or '393' in t or 'Your Name' in t:
            _fill_runs(p, exact_map={'Planner Name': name, 'Planner': name},
                       contains_map={'Your Name at 832-393-XXXX': f"{name} at {phone}", '832-393-XXXX': phone,
                                     '(832) 393-XXXX': phone, 'Planner Name': name, 'Your Name': name,
                                     'or call the ': 'or call '})
            if not _has_hyperlink(p):
                _replace_in_runs(p, {'Planner Name at 832-393 XXX': f"{name} at {phone}",
                                     **{k: email for k in _EMAIL_PLACEHOLDERS}})
            _fill_phone(p, phone)
        _replace_in_runs(p, common)
    _replace_text_nodes(doc, {k: email for k in _EMAIL_PLACEHOLDERS})
    _fix_mailto_links(doc, email)


def generate_notification(workbook_path, template_path, map_path, output_path, mode='preview'):
    ctx = _get_context(workbook_path)
    owners = _get_owners(load_workbook(workbook_path, data_only=True))
    if not owners:
        raise RuntimeError('No valid owner rows found on Step 1.')
    selected = owners[:1] if mode == 'preview' else owners
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        map_png = _map_to_png(map_path, td) if map_path else None
        base = None
        for owner in selected:
            doc = Document(template_path)
            _fill_owner_letter(doc, ctx, owner)
            _compact_notification_first_page(doc)
            _insert_map(doc, map_png)
            if base is None:
                base = doc
            else:
                _restart_list_numbering(base, doc)
                _append_letter_as_own_section(base, doc)
        base.save(str(out))
    unlock_docx(out)
    return str(out)


# =============================================================================
# Applicant sign letter
# =============================================================================
def _applicant_common(ctx):
    return {
        '«Owner_Mail»': ctx['Applicant_Name'],
        '«Mail_Addr_»': ctx['Applicant_Address_Block'],
        '«Mail_City»': ctx['Applicant_City'] + ',',
        '«Mail_State»': ctx['Applicant_State'],
        '«Mail_Zip»': ctx['Applicant_Zip'],
        'XXX block of XX Street, XX and XX sides,': ctx['Area_Line_1'] or 'XXX block of XX Street, XX and XX sides,',
        'between XX and XX': ctx['Area_Line_2'] or 'between XX and XX',
    }


def generate_applicant_letter(workbook_path, template_path, output_path):
    ctx = _get_context(workbook_path)
    doc = Document(template_path)
    _unwrap_smart_tags(doc)
    common = _applicant_common(ctx)
    for p in doc.paragraphs:
        t = _norm(p.text)
        ts = t.strip()
        if ts == 'DATE':
            _set_exact_para_preserve_first_run(p, ctx['Letter_Date'])
            continue
        if ts.startswith('Dear'):
            _fill_runs(p, exact_map={'Property Owners Name': ctx['Applicant_Name']})
        elif 'Minimum Building Line is' in t:
            _fill_runs(p, exact_map={
                'X': ctx['Minimum_Building_Line_Text'], 'XX': ctx['Minimum_Building_Line_Text'],
                'two (2': f"{ctx['Sign_Count_Words']} ({ctx['Sign_Count_Number']}",
                'one': ctx['Sign_Placement']}, contains_map={' sign at each corner.': '.'})
        elif 'feel free to contact' in t:
            _fill_runs(p, exact_map={'Planner': ctx['Planner_Name'], 'Planner Name': ctx['Planner_Name'], 'xxxx': '', 'XXXX': '',
                                     'Planner Name@houstontx.gov.': ctx['Planner_Email'] + '.'},
                       contains_map={' at (832) 393-': f" at {ctx['Planner_Phone']}"})
        _replace_in_runs(p, common)
    _replace_text_nodes(doc, {k: ctx['Planner_Email'] for k in _EMAIL_PLACEHOLDERS})
    _fix_mailto_links(doc, ctx['Planner_Email'])
    _clear_all_red_run_colors(doc)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    unlock_docx(out)
    return 1


# =============================================================================
# Approvals
# =============================================================================
def _replace_line_runs(p, marker, new_text):
    """Replace the <w:br/>-separated line that contains `marker` inside a
    criteria cell, keeping that line's own run formatting (the City bolds the
    finding lines while the criterion above them is italic)."""
    lines, cur = [], []
    for r in p.runs:
        if r._r.find(qn('w:br')) is not None and not (r.text or '').strip():
            lines.append(cur)
            cur = []
            continue
        cur.append(r)
    lines.append(cur)
    for runs in lines:
        joined = _norm(''.join(r.text for r in runs))
        if marker in joined and runs:
            target = next((r for r in runs if (r.text or '').strip()), runs[0])
            for r in runs:
                _set_run_text_preserve_xml(r, '')
            _set_run_text_preserve_xml(target, new_text)
            _blacken_run(target)
            return True
    return False


def _set_line(p, marker, new_text):
    if not _replace_line_runs(p, marker, new_text):
        _replace_line_in_paragraph(p, marker, new_text)


def generate_director_approval(workbook_path, template_path, output_path):
    ctx = _get_context(workbook_path)
    from core.city_templates import render_updated
    if render_updated(template_path, output_path, ctx, workbook_path, 'mbl', 'director'): return 1
    doc = Document(template_path)
    _unwrap_smart_tags(doc)
    for p in _iter_all_paragraphs(doc):
        t = _norm(p.text)
        ts = t.strip()
        if 'Block No.' in t:
            _replace_runs_contextual(p, exact_map={'XX': ctx['Application_Number'], 'XXX': ctx['Application_Number']})
        elif 'minimum building line of' in t:
            _replace_runs_contextual(p, exact_map={'XXXX': ctx['Minimum_Building_Line_Text'], 'XX': ctx['Minimum_Building_Line_Text']})
        elif 'The application is for the' in t:
            _set_line(p, 'The application is for the', f"The application is for the {ctx['Area_Description']}.")
        elif ts == 'Vonn Tran':
            _set_exact_para_preserve_first_run(p, ctx['Director_Name'])
        elif ts == 'Date':
            _set_exact_para_preserve_first_run(p, _clean(ctx.get('Director_Approval_Date')) or 'Date')
    _clear_all_red_run_colors(doc)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    unlock_docx(out)
    return 1


def generate_pc_approval(workbook_path, template_path, output_path):
    ctx = _get_context(workbook_path)
    from core.city_templates import render_updated
    if render_updated(template_path, output_path, ctx, workbook_path, 'mbl', 'pc'): return 1
    doc = Document(template_path)
    _unwrap_smart_tags(doc)
    app = ctx['Application_Number']
    min_bl = ctx['Minimum_Building_Line_Text']
    qual = _to_int(ctx.get('Qualifying_Lot_Count'))
    for p in _iter_all_paragraphs(doc):
        t = _norm(p.text)
        ts = t.strip()
        if 'Requirement Area No.' in t or 'Block No.' in t:
            if not _replace_runs_contextual(p, exact_map={'106': app, 'XX': app, 'XXX': app}):
                c = _RunCursor(p)
                if not c.span('10', '6', app):
                    _replace_in_runs(p, {'106': app})
        elif 'The application is for the' in t:
            _set_line(p, 'The application is for the', f"The application is for the {ctx['Area_Description']}.")
        elif '% of the proposed application area' in t:
            _set_line(p, '% of the proposed application area',
                      f"{ctx['SFR_Percentage_Number']}% of the proposed application area is developed with single-family residential units per property.")
        elif 'Petition signed by owners of' in t:
            _set_line(p, 'Petition signed by owners of', f"Petition signed by owners of {ctx['Support_Percentage_Number']}% of the SMBLB.")
        elif 'A minimum building line of' in t:
            _set_line(p, 'A minimum building line of',
                      f"A minimum building line of {min_bl} feet exists on {_num_words(qual) if qual else 'X'} ({qual if qual else 'X'}) properties in the blockface.")
        elif 'The subdivision was platted in' in t:
            _set_exact_para_preserve_first_run(p, ctx['PC_Narrative'])
        elif 'developed properties (representing' in t:
            _set_exact_para_preserve_first_run(p, ctx['Qualifying_Sentence'])
        elif ts == 'Date':
            _set_exact_para_preserve_first_run(p, _clean(ctx.get('PC_Approval_Date')) or 'Date')
    _clear_all_red_run_colors(doc)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    unlock_docx(out)
    return 1


# =============================================================================
# Routing form (2026 coversheet routing form, content controls)
# =============================================================================
def _drop_blank_final_section(doc):
    """The City routing form ends with an empty paragraph after the two-column
    attachments section, which spills onto a blank second page."""
    body = doc.element.body
    kids = list(body)
    while len(kids) >= 2 and kids[-2].tag == qn('w:p') and not ''.join(t.text or '' for t in kids[-2].iter(qn('w:t'))).strip() \
            and kids[-2].find(qn('w:pPr') + '/' + qn('w:sectPr')) is None and not list(kids[-2].iter(qn('w:drawing'))):
        body.remove(kids[-2])
        kids = list(body)
    if len(kids) >= 2 and kids[-1].tag == qn('w:sectPr') and kids[-2].tag == qn('w:p'):
        ppr = kids[-2].find(qn('w:pPr'))
        psect = ppr.find(qn('w:sectPr')) if ppr is not None else None
        if psect is not None:
            final = kids[-1]
            for tag in ('w:pgSz', 'w:pgMar', 'w:docGrid'):
                if psect.find(qn(tag)) is None and final.find(qn(tag)) is not None:
                    psect.append(copy.deepcopy(final.find(qn(tag))))
            ppr.remove(psect)
            body.remove(final)
            body.append(psect)


def generate_routing_form(workbook_path, template_path, output_path):
    ctx = _get_context(workbook_path)
    from core.city_templates import render_updated
    if render_updated(template_path, output_path, ctx, workbook_path, 'mbl', 'routing'): return 1
    doc = Document(template_path)
    from core.mbl_templates_2026 import fill_mbl_template
    if fill_mbl_template(doc, ctx, renewal=False, routing=True):
        _clear_all_red_run_colors(doc)
        prepare_routing_form(doc)
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out))
        unlock_docx(out)
        return 1
    is_pc = ctx['Approval_Path'] == 'Planning Commission Approval'
    app = ctx['Application_Number']
    location = f"The {ctx['Area_Description']}" if ctx.get('Block_Description') else 'The X block of X Street, X and X sides, between X Street and X Street'
    paras = doc.paragraphs
    for i, p in enumerate(paras):
        t = _norm(p.text)
        if t.startswith('The topic of Request'):
            _fill_sdts_in_order(p, [f"Special Minimum Building Line Block Application No. {app}"])
            if i + 1 < len(paras) and paras[i + 1].text.strip().startswith('('):
                loc = paras[i + 1]
                _fill_sdts_in_order(loc, [location])
                ind = loc._p.get_or_add_pPr().find(qn('w:ind'))
                if ind is not None:
                    left = int(ind.get(qn('w:left'), '0') or 0) + int(ind.get(qn('w:firstLine'), '0') or 0)
                    ind.set(qn('w:left'), str(left))
                    ind.set(qn('w:firstLine'), '0')
                    if qn('w:hanging') in ind.attrib:
                        del ind.attrib[qn('w:hanging')]
        elif t.startswith('Coversheet Author'):
            _fill_sdts_in_order(p, [f"{ctx['Planner_Name']}, {ctx['Planner_Phone']}"])
        elif 'no later than' in t:
            _fill_sdts_in_order(p, [ctx['Council_Deadline'] or 'DATE'])
        elif 'BUILDING LINE' in t and 'LOT SIZE' in t:
            _set_paragraph_checkboxes(p, [True, False])
            _fill_sdts_in_order(p, [app, f": {ctx['Minimum_Building_Line_Text']}", ctx['SFR_Percentage_Number']])
        elif t.startswith('Petition Signed By'):
            _fill_sdts_in_order(p, [f" {ctx['Support_Percentage_Number']}", f"{ctx['Total_Lot_Count']} / {ctx['Qualifying_Lot_Count']}"])
        elif 'SMLS Area Application' in t:
            _set_paragraph_checkboxes(p, [False, False, True])
        elif 'SMLS Block Renewal' in t and 'SMBL' not in t:
            _set_paragraph_checkboxes(p, [False])
        elif 'SMBL Block Renewal' in t:
            _set_paragraph_checkboxes(p, [False])
        elif 'Planning Commission Approval' in t and 'Director' in t:
            _set_paragraph_checkboxes(p, [is_pc, not is_pc, True, True, True, True, True, False])
    _clear_all_red_run_colors(doc)
    _drop_blank_final_section(doc)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    prepare_routing_form(doc)
    doc.save(str(out))
    unlock_docx(out)
    return 1


# =============================================================================
# OneMeeting staff report (Director / support under 51% / protest)
# =============================================================================
def _set_segments(p, segments):
    """Rebuild a paragraph from (text, bold) segments, cloning the run
    formatting of the template paragraph."""
    runs = list(p.runs)
    rpr = None
    for r in runs:
        if r._r.find(qn('w:rPr')) is not None:
            rpr = deepcopy(r._r.find(qn('w:rPr')))
            break
    for r in runs:
        r._r.getparent().remove(r._r)
    for text, bold in segments:
        if not text:
            continue
        run = p.add_run()
        if rpr is not None:
            run._r.insert(0, deepcopy(rpr))
        run.text = text
        if bold is not None:
            run.bold = bold
        _blacken_run(run)


def staff_report_variant(ctx):
    return ctx.get('Staff_Report_Variant', 'director')


def generate_staff_report(workbook_path, template_path, output_path):
    ctx = _get_context(workbook_path)
    from core.city_templates import render_updated
    if render_updated(template_path, output_path, ctx, workbook_path, 'mbl', 'staff'): return 1
    doc = Document(template_path)
    _unwrap_smart_tags(doc)
    full = _norm(' '.join(p.text for p in doc.paragraphs[:3])).upper()
    variant = 'less51' if 'LESS THAN 51' in full else ('protest' if 'PROTEST' in full else 'director')
    area = ctx['Area_Description'] or 'XXXX block of XXXX Street, XX side, between XX Street and XX Street'
    caption = ctx['Ordinance_Caption'] or 'XX side of the XXXX Blocks of XX Street, between XX Street and XX Street'
    district = ctx['Council_District'] or 'X'
    subdivision = ctx['Subdivision_Name']
    if subdivision != 'SUBDIVISION NAME' and 'subdivision' not in subdivision.lower():
        subdivision = f"{subdivision} Subdivision"
    min_bl = ctx['Minimum_Building_Line_Text']
    hearing = ctx['PC_Hearing_Date'] if ctx['PC_Hearing_Date'] != '[Date]' else '(DATE OF PUBLIC HEARING)'
    n = ctx['Protest_Count']
    bold_values = variant == 'director'
    remove = []
    paras = doc.paragraphs
    for i, p in enumerate(paras):
        t = _norm(p.text)
        c = _RunCursor(p)
        if t.startswith('DISTRICT:'):
            if ctx['Council_District']:
                _replace_in_runs(p, {'CHOOSE THE DISTRICT': ctx['Council_District']})
        elif 'Block Application No.' in t:
            _replace_runs_contextual(p, exact_map={'XXX': ctx['Application_Number']})
        elif t.startswith('(') and 'block of' in t:
            _replace_in_runs(p, {'2800 - 2900 block of Wheeler Street, north side, between Delano Street and Ennis Street': area,
                                 'XXXX block of XXXX Street, XX side, between XX Street and XX Street': area})
        elif t.startswith('ITEM CAPTION'):
            c.set('ENTER THE CAPTION FROM THE ORDINANCE (see below example)', '')
            if not c.set('the north side of the 2800-2900 Blocks of Wheeler Street, between Delano Street and Ennis Street ', f"the {caption} "):
                c.span('the ', 'Street ', f"the {caption} ")
            if not c.set('DISTRICT D - EVANS-SHABAZZ', f"DISTRICT {district} - {ctx['Council_Member']}"):
                c.span('DISTRICT ', 'COUNCIL MEMBER NAME', f"DISTRICT {district} - {ctx['Council_Member']}")
        elif t.startswith('In accordance with'):
            head = [("In accordance with Section 42-170 of the Code of Ordinances, the property owner at ", False),
                    (ctx['Applicant_Property_Address'], bold_values), (" of the ", False), (subdivision, bold_values),
                    (" submitted an application to designate the ", False), (area, bold_values),
                    (" as a Special Minimum Building Line Block (SMBLB). ", False)]
            if variant == 'director':
                segs = head + [("The application includes support from ", False), (f"{ctx['Support_Percentage_Number']}%", True),
                               (" of property owners on the block. Notifications were mailed to all property owners within the boundary, advising them of the application and the 30-day protest period. ", False),
                               (_protest_sentence(n, 'No protests were received.'), False)]
                if i + 1 < len(paras) and 'as a Special Minimum Building Line Block' in _norm(paras[i + 1].text) and not _norm(paras[i + 1].text).startswith('In accordance'):
                    remove.append(paras[i + 1])
            elif variant == 'less51':
                segs = head + [("Notifications were mailed to all property owners within the boundary, advising them of the application and the 30-day protest period. ", False),
                               (_protest_sentence(n, 'No protest was filed.'), False)]
            else:
                segs = head + [("Notifications were mailed to all property owners within the boundary, advising them of the application and the 30-day protest period. ", False),
                               (_protest_sentence(n, 'No protests were received.'), False),
                               ("  The Houston Planning Commission considered the protested application on ", False), (hearing, False),
                               (" and voted to recommend that the City Council establish the SMBLB.", False)]
            _set_segments(p, segs)
        elif t.startswith('As no action is required'):
            _set_segments(p, [("As no action is required by the Houston Planning Commission, the Planning and Development Department is forwarding the application to City Council with a recommendation to adopt an ordinance establishing a Special Minimum Building Line of ", False),
                              (f"{min_bl} ", True), ("feet for the ", False), (area, True), (". ", False)])
        elif t.startswith('This application had less than 51% support'):
            _set_segments(p, [("This application had less than 51% support and was referred to the Houston Planning Commission for review and consideration. The Houston Planning Commission considered the application on ", False),
                              (hearing, True), (" and voted to recommend that the City Council establish the SMBLB. The Planning and Development Department is forwarding the application to City Council with a recommendation to adopt an ordinance establishing a Special Minimum Building Line of ", False),
                              (f"{min_bl} feet", False), (" for the ", False), (area, False), (".  ", False)])
        elif t.startswith('The Planning and Development Department is forwarding'):
            _set_segments(p, [("The Planning and Development Department is forwarding the application to City Council with a recommendation to adopt an ordinance establishing a Special Minimum Building Line of ", False),
                              (f"{min_bl} feet", False), (" for the ", False), (area, False), (".  ", False)])
        elif 'Your Name, Position, Phone' in t or 'Your Name, Planner, Phone' in t:
            contact = f"{ctx['Planner_Name']}, {ctx['Planner_Position']}, {ctx['Planner_Phone']}"
            _replace_in_runs(p, {'Your Name, Position, Phone Number': contact, 'Your Name, Planner, Phone Number': contact})
    for p in remove:
        p._p.getparent().remove(p._p)
    _clear_all_red_run_colors(doc)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    unlock_docx(out)
    return 1


# =============================================================================
# Draft ordinance (token engine until the City provides an unlocked 2026 file)
# =============================================================================
def generate_draft_ordinance(workbook_path, template_path, map_path, output_path):
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    return _legacy_generate_draft_ordinance(str(workbook_path), str(template_path), str(map_path), str(out))


# =============================================================================
# RCA packet, HCAD
# =============================================================================
def _split_paths(raw):
    if not raw:
        return []
    if isinstance(raw, str):
        return [x for x in re.split(r'[;|]+', raw) if x.strip()]
    return list(raw)


def generate_rca_packet_pdf(workbook_path, templates, map_path, subdivision_plat_path, tax_map_path,
                            hcad_records_path, output_pdf, include_staff_analysis=True):
    """Routing form, Director or PC approval, boundary map, Step 2 staff
    analysis, [plat], tax map, [HCAD records condensed to one page each]."""
    ctx = _get_context(workbook_path)
    output_pdf = Path(output_pdf)
    from core.packet_work import packet_workspace
    with packet_workspace() as work:
        for label, val in (('boundary map', map_path), ('tax map', tax_map_path)):
            if not str(val or '').strip():
                raise ValueError(f'Select the {label} before building the RCA packet.')
        routing = work / '01_MBL_Routing_Form.docx'
        generate_routing_form(workbook_path, templates['routing'], routing)
        pdfs = [routing_pdf_one_page(_ensure_pdf(routing, work), work)]
        if ctx['Approval_Path'] == 'Planning Commission Approval':
            pc = work / '02_MBL_PC_Approval.docx'
            generate_pc_approval(workbook_path, templates['pc'], pc)
            pdfs.append(_ensure_pdf(pc, work))
        else:
            da = work / '02_MBL_Director_Approval.docx'
            generate_director_approval(workbook_path, templates['director'], da)
            pdfs.append(_ensure_pdf(da, work))
        pdfs.append(_ensure_pdf(map_path, work))
        if include_staff_analysis:
            staff = work / '04_MBL_Staff_Analysis_Step2.pdf'
            export_mbl_step2_staff_analysis_pdf(workbook_path, staff)
            pdfs.append(staff)
        if str(subdivision_plat_path or '').strip():
            pdfs.append(_ensure_pdf(subdivision_plat_path, work))
        pdfs.append(_ensure_pdf(tax_map_path, work))
        hcad = _split_paths(hcad_records_path)
        if hcad:
            pdfs.extend(condense_hcad_pdfs([_ensure_pdf(p, work) for p in hcad], work))
        _merge_pdfs(pdfs, output_pdf)
        return len(pdfs)


# =============================================================================
# UI
# =============================================================================
_TEMPLATE_ROWS = [
    ('Owner notification letter', 'template', 'MBL_ALL_OWNERS.docx'),
    ('Notice of public hearing (support under 51%)', 'hearing_template', 'MBL_Notice_Public_Hearing_Less_51.docx'),
    ('Protest letter to owners', 'protest_template', 'MBL_Protest_All_Owners.docx'),
    ('Notice of PC rescheduling', 'reschedule_template', 'MBL_Notice_PC_Rescheduling.docx'),
    ('Applicant sign letter', 'applicant_template', 'MBL_APPLICANT_COMPLETE_SIGNS.docx'),
    ('Director approval', 'director_template', 'MBL_Directors_Approval.docx'),
    ('Planning Commission approval', 'pc_template', 'MBL_PC_Approval.docx'),
    ('Routing form', 'routing_template', 'MBL_Routing_Form.docx'),
    ('OneMeeting staff report (Director)', 'staff_report_template', 'MBL_OneMeeting_Staff_Report.docx'),
    ('OneMeeting staff report (under 51%)', 'staff_report_less51_template', 'MBL_OneMeeting_Staff_Report_Less51.docx'),
    ('OneMeeting staff report (protest)', 'staff_report_protest_template', 'MBL_OneMeeting_Staff_Report_Protest.docx'),
    ('MBL draft ordinance', 'ordinance_template', 'MBL_Ordinance_Template_Clean.docx'),
]
_STAFF_KEYS = {'director': 'staff_report_template', 'less51': 'staff_report_less51_template', 'protest': 'staff_report_protest_template'}


class MBLWizard(EmbeddedWizardMixin, tk.Frame):
    _boundary_map_wizard_key = _WIZARD_KEY
    def __init__(self, parent=None):
        super().__init__(parent)
        self.title(f'MBL Planner Wizard {WIZARD_VERSION}')
        self.geometry('1280x820')
        base = sample_input_dir(_WIZARD_KEY)
        self.files = {
            'workbook': str(calculation_template_path('mbl')),
            'map': str(base / 'sample_boundary_map.pdf'),
            'tax_map': '',
            'subdivision_plat': '',
            'hcad_records': '',
            'include_staff_analysis': 'yes',
            'out': str(Path.home() / 'Documents'),
        }
        for _lab, key, fn in _TEMPLATE_ROWS:
            self.files[key] = str(template_path(_WIZARD_KEY, key, fn))
        for _k, _v in core_session.load(_WIZARD_KEY).items():
            if _k in self.files and not self._is_foreign_bundled_template(_k, _v):
                self.files[_k] = _v
        self._build()

    def _is_foreign_bundled_template(self, key, saved_path):
        """A path saved by an older suite install (another templates\\mbl folder)
        must not override the bundled 2026 template."""
        if not (key == 'template' or key.endswith('_template')) or not saved_path:
            return False
        try:
            saved = Path(saved_path)
            default = Path(self.files[key])
            if saved.resolve() == default.resolve():
                return False
            return saved.parent.name.lower() == 'mbl' and saved.parent.parent.name.lower() == 'templates'
        except Exception:
            return False

    def _build(self):
        self.colors = {'nav': '#14532D', 'nav2': '#166534', 'nav_active': '#16A34A', 'bg': '#F8FAFC', 'card': '#FFFFFF',
                       'border': '#CBD5E1', 'muted': '#64748B', 'text': '#0F172A', 'accent_light': '#DCFCE7', 'accent_mid': '#86EFAC'}
        self.configure(bg=self.colors['bg'])
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        sidebar = tk.Frame(self, bg=self.colors['nav'], width=255)
        sidebar.grid(row=0, column=0, sticky='ns')
        sidebar.grid_propagate(False)
        tk.Label(sidebar, text=f'MBL\nWizard {WIZARD_VERSION}', bg=self.colors['nav'], fg='white', font=('Segoe UI', 20, 'bold'), justify='left').pack(anchor='nw', padx=22, pady=(24, 8))
        tk.Label(sidebar, text='Document automation for Special Minimum Building Line Block RCA packets (2026 City templates)', bg=self.colors['nav'], fg='#DCFCE7', font=('Segoe UI', 9), justify='left', wraplength=200).pack(anchor='nw', padx=22, pady=(0, 24))
        self.step_buttons = []
        steps = [('1. Select Files', self._select_files_page), ('2. Review Case Info', self._case_info_page), ('3. Review Owners', self._owners_page), ('4. Generate Documents', self._generate_page), ('5. Peer Reviewer', self._peer_review_page)]
        for i, (label, cmd) in enumerate(steps):
            b = tk.Button(sidebar, text=label, anchor='w', bg=self.colors['nav_active'] if i == 0 else self.colors['nav2'], fg='white', activebackground=self.colors['nav_active'], activeforeground='white', relief='flat', font=('Segoe UI', 10, 'bold' if i == 0 else 'normal'), padx=14, pady=12, command=cmd)
            b.pack(fill='x', padx=12, pady=5)
            self.step_buttons.append(b)
        modules = ('Built modules\n• Case Info (2026 workbook)\n• Owner letters\n• Hearing / protest / rescheduling letters\n• Applicant sign letter\n'
                   '• Director / PC approvals\n• 2026 routing form\n• OneMeeting staff reports\n• Draft ordinance\n• RCA packet PDF')
        tk.Label(sidebar, text=modules, bg=self.colors['nav'], fg='#BBF7D0', font=('Segoe UI', 9), justify='left').pack(side='bottom', anchor='sw', padx=22, pady=22)
        main = tk.Frame(self, bg=self.colors['bg'])
        main.grid(row=0, column=1, sticky='nsew')
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)
        header = tk.Frame(main, bg=self.colors['bg'])
        header.grid(row=0, column=0, sticky='ew', padx=22, pady=(18, 14))
        header.grid_columnconfigure(0, weight=1)
        tk.Label(header, text='MBL/RCA Document Generator', bg=self.colors['bg'], fg=self.colors['text'], font=('Segoe UI', 25, 'bold')).grid(row=0, column=0, sticky='w')
        tk.Label(header, text='Load one MBL calculation workbook, review Case Info, then generate the City templates and the RCA packet.', bg=self.colors['bg'], fg='#334155', font=('Segoe UI', 12), wraplength=950, justify='left').grid(row=1, column=0, sticky='w', pady=(4, 0))
        tk.Button(header, text='Load Workbook', command=self._case_info_page, bg=self.colors['accent_light'], activebackground=self.colors['accent_mid'], fg=self.colors['text'], font=('Segoe UI', 10, 'bold'), padx=16, pady=10, relief='solid', bd=1).grid(row=0, column=1, rowspan=2, padx=(14, 0), sticky='e')
        _holder = ScrollableContent(main, bg=self.colors['bg'])
        _holder.grid(row=1, column=0, sticky='nsew', padx=22)
        self.content = _holder.interior
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)
        status = tk.LabelFrame(main, text='Status', bg=self.colors['card'], fg=self.colors['text'], font=('Segoe UI', 11, 'bold'), padx=14, pady=10)
        status.grid(row=2, column=0, sticky='ew', padx=22, pady=(8, 14))
        self.status_var = tk.StringVar(value=f'Ready. MBL Wizard {WIZARD_VERSION} loaded with the 2026 City templates.')
        tk.Label(status, textvariable=self.status_var, bg=self.colors['card'], fg='#475569', font=('Segoe UI', 10), anchor='w', justify='left').pack(fill='x')
        self._select_files_page()

    def clear(self):
        for w in self.content.winfo_children():
            w.destroy()

    def _set_active(self, idx):
        for i, b in enumerate(self.step_buttons):
            b.configure(bg=self.colors['nav_active'] if i == idx else self.colors['nav2'], font=('Segoe UI', 10, 'bold' if i == idx else 'normal'))

    def _section_card(self, parent, title, row, col=0, colspan=1, pady=8):
        f = tk.LabelFrame(parent, text=title, bg='white', fg=self.colors['text'], font=('Segoe UI', 11, 'bold'), padx=12, pady=10)
        f.grid(row=row, column=col, columnspan=colspan, sticky='nsew', padx=8, pady=pady)
        f.grid_columnconfigure(1, weight=1)
        return f

    def _file_row(self, parent, label, key, row, folder=False, multi=False):
        tk.Label(parent, text=label, bg='white', fg='#334155', font=('Segoe UI', 9, 'bold'), anchor='w').grid(row=row, column=0, sticky='w', padx=10, pady=8)
        path_var = tk.StringVar(master=parent, value=self.files.get(key, ''))
        ent = tk.Entry(parent, textvariable=path_var, width=62, font=('Segoe UI', 9))
        ent.grid(row=row, column=1, sticky='ew', padx=8, pady=8)
        # Keep typed/pasted edits, clears and Browse selections in the same state.
        ent._path_variable = path_var
        path_var.trace_add('write', lambda *_: self.files.__setitem__(key, path_var.get()))

        def browse():
            if folder:
                p = filedialog.askdirectory()
            elif multi:
                p = ';'.join(filedialog.askopenfilenames(filetypes=[('PDF files', '*.pdf'), ('All files', '*.*')]))
            else:
                p = filedialog.askopenfilename(filetypes=[('Supported files', '*.xlsx *.xls *.docx *.doc *.pdf *.png *.jpg *.jpeg'), ('All files', '*.*')])
            if p:
                self.files[key] = p
                ent.delete(0, 'end')
                ent.insert(0, p)
        tk.Button(parent, text='Browse Multiple' if multi else 'Browse', command=browse, bg='#E7E5E4', font=('Segoe UI', 9), padx=14, pady=7, relief='solid', bd=1).grid(row=row, column=2, sticky='e', padx=10, pady=8)

    def _select_files_page(self):
        self._set_active(0)
        self.clear()
        outer = tk.Frame(self.content, bg=self.colors['bg'])
        outer.grid(row=0, column=0, sticky='nsew')
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_columnconfigure(1, weight=1)
        tk.Label(outer, text='Step 1: Select the files for this case', bg=self.colors['bg'], fg=self.colors['text'], font=('Segoe UI', 16, 'bold')).grid(row=0, column=0, columnspan=2, sticky='w', padx=8, pady=(0, 6))
        tk.Label(outer, text='City templates are prefilled from the templates folder. Use Browse only if the City updates a template.', bg=self.colors['bg'], fg=self.colors['muted'], font=('Segoe UI', 10), wraplength=980, justify='left').grid(row=1, column=0, columnspan=2, sticky='w', padx=8, pady=(0, 10))
        setup = self._section_card(outer, 'Required setup', 2, 0, 2)
        self._file_row(setup, 'MBL calculation workbook', 'workbook', 0)
        self._file_row(setup, 'Boundary map', 'map', 1)
        self._file_row(setup, 'Output folder', 'out', 2, folder=True)
        self._add_boundary_map_controls(setup)
        self._add_application_import_controls(setup)
        templates = self._section_card(outer, 'City templates (2026)', 3, 0)
        for i, (lab, key, _fn) in enumerate(_TEMPLATE_ROWS):
            self._file_row(templates, lab, key, i)
        att = self._section_card(outer, 'Packet attachments', 3, 1)
        tk.Label(att, text='Upload only what the wizard cannot create: the tax map (required), the subdivision plat and HCAD records (optional). The Step 2 staff analysis is exported by the wizard.', bg='white', fg=self.colors['muted'], font=('Segoe UI', 9), wraplength=390, justify='left').grid(row=0, column=0, columnspan=3, sticky='w', padx=10, pady=(0, 6))
        self._file_row(att, 'Tax map (required)', 'tax_map', 1)
        self._file_row(att, 'Subdivision plat (optional)', 'subdivision_plat', 2)
        self._file_row(att, 'HCAD records (optional)', 'hcad_records', 3, multi=True)
        self.staff_var = tk.BooleanVar(value=str(self.files.get('include_staff_analysis', 'yes')).lower() in ('yes', 'true', '1'))
        tk.Checkbutton(att, text='Include the Step 2 staff analysis export in the packet', variable=self.staff_var,
                       command=lambda: self.files.__setitem__('include_staff_analysis', 'yes' if self.staff_var.get() else 'no'),
                       bg='white', fg='#334155', font=('Segoe UI', 9), anchor='w').grid(row=4, column=0, columnspan=3, sticky='w', padx=10, pady=(2, 4))
        hcad_btns = tk.Frame(att, bg='white')
        hcad_btns.grid(row=5, column=0, columnspan=3, sticky='w', padx=10, pady=(2, 4))
        tk.Button(hcad_btns, text='Download HCAD Records', command=self._download_hcad_records, bg=self.colors['accent_light'], activebackground=self.colors['accent_mid'], font=('Segoe UI', 9, 'bold'), padx=12, pady=6, relief='solid', bd=1).pack(side='left', padx=(0, 6))
        tk.Button(hcad_btns, text='Import Downloaded PDFs', command=self._import_hcad_pdfs, bg='#E7E5E4', font=('Segoe UI', 9), padx=12, pady=6).pack(side='left', padx=(0, 6))
        tk.Button(hcad_btns, text='Open HCAD Folder', command=self._open_hcad_folder, bg='#E7E5E4', font=('Segoe UI', 9), padx=12, pady=6).pack(side='left')
        tk.Label(att, text='HCAD: click Download HCAD Records to read Step 1, open Google Chrome for HCAD verification, then save the HCAD PDFs into the HCAD_Records folder inside your output folder and select them for the packet. If HCAD blocks automation, save PDFs there and click Import Downloaded PDFs.', bg='white', fg=self.colors['muted'], font=('Segoe UI', 9), wraplength=390, justify='left').grid(row=6, column=0, columnspan=3, sticky='w', padx=10, pady=(0, 8))

    # ---------------- HCAD download workflow ----------------
    def _hcad_folder(self):
        return Path(self.files.get('out') or (Path.home() / 'Documents')) / 'HCAD_Records'

    def _chrome_debug_profile(self):
        return self._hcad_folder() / 'Chrome_Debug_Profile'

    def _chrome_debug_url(self):
        return 'http://127.0.0.1:9222'

    def _is_chrome_cdp_available(self):
        try:
            with urllib.request.urlopen(self._chrome_debug_url() + '/json/version', timeout=2) as resp:
                return resp.status == 200
        except Exception:
            return False

    def _find_chrome_executable(self):
        candidates = []
        for env in ('LOCALAPPDATA', 'PROGRAMFILES', 'PROGRAMFILES(X86)'):
            v = os.environ.get(env)
            if v:
                candidates.append(Path(v) / 'Google' / 'Chrome' / 'Application' / 'chrome.exe')
        for c in candidates:
            if c.exists():
                return str(c)
        return 'chrome'

    def _export_hcad_list(self, show_message=True):
        if not self.files.get('workbook') or not self.files.get('out'):
            messagebox.showerror('Missing files', 'Select the MBL workbook and output folder first.')
            return None, 0
        try:
            address_file, count = export_hcad_address_list(self._wb(), self.files['out'])
            if show_message:
                messagebox.showinfo('HCAD', f'Prepared {count} unique HCAD property record(s). One record is used for all lot rows sharing an HCAD account.\n\n{address_file}')
            return address_file, count
        except Exception as e:
            messagebox.showerror('HCAD address export error', str(e))
            return None, 0

    def _start_hcad_chrome_session(self, show_message=True):
        hcad_folder = self._hcad_folder()
        hcad_folder.mkdir(parents=True, exist_ok=True)
        profile = self._chrome_debug_profile()
        profile.mkdir(parents=True, exist_ok=True)
        try:
            subprocess.Popen([self._find_chrome_executable(), '--remote-debugging-port=9222', f'--user-data-dir={profile}', '--new-window', 'https://search.hcad.org/'],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            for _ in range(20):
                if self._is_chrome_cdp_available():
                    break
                time.sleep(0.25)
            if show_message:
                messagebox.showinfo('HCAD', 'Google Chrome was opened for HCAD.\n\n1. Complete the HCAD human verification in that Chrome window.\n2. Wait until the HCAD property search page is visible.\n3. Return to this wizard and click OK/Continue.')
            return True
        except Exception as e:
            messagebox.showerror('HCAD', f'Could not start Google Chrome for HCAD:\n{e}\n\nMake sure Google Chrome is installed.')
            return False

    def _import_hcad_pdfs(self):
        folder = self._hcad_folder()
        folder.mkdir(parents=True, exist_ok=True)
        pdfs = sorted(str(p) for p in folder.glob('*.pdf') if p.is_file())
        if not pdfs:
            messagebox.showinfo('HCAD', f'No HCAD PDFs were found in:\n{folder}\n\nSave downloaded HCAD PDFs into this folder, then click Import Downloaded PDFs again.')
            return []
        self.files['hcad_records'] = ';'.join(pdfs)
        messagebox.showinfo('HCAD', f'Imported {len(pdfs)} HCAD PDF(s) from {folder}.')
        self._select_files_page()
        return pdfs

    def _open_hcad_folder(self):
        folder = self._hcad_folder()
        folder.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform.startswith('win'):
                os.startfile(str(folder))
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', str(folder)])
            else:
                subprocess.Popen(['xdg-open', str(folder)])
        except Exception as e:
            messagebox.showerror('HCAD', f'Could not open HCAD folder:\n{e}')

    def _download_hcad_records(self):
        address_file, count = self._export_hcad_list(show_message=False)
        if not address_file:
            return
        if count <= 0:
            messagebox.showwarning('HCAD', 'No HCAD addresses or TAX_ID values were found in Step 1.')
            return
        hcad_folder = self._hcad_folder()
        hcad_folder.mkdir(parents=True, exist_ok=True)
        script_path = HCAD_DOWNLOADER
        if not core_launch.hcad_available():
            messagebox.showerror('HCAD', f'Could not find bundled HCAD downloader:\n{script_path}')
            return
        if not self._is_chrome_cdp_available():
            if not self._start_hcad_chrome_session(show_message=False):
                return
        messagebox.showinfo('HCAD', f'HCAD will open in Google Chrome for {count} record(s).\n\nComplete the human verification in Chrome if prompted. Do not close the Chrome window.\n\nWhen the HCAD search page is visible, click OK here and the wizard will continue downloading to the HCAD_Records folder.')
        if not self._is_chrome_cdp_available():
            messagebox.showwarning('HCAD', 'The wizard still cannot connect to Chrome.\n\nMake sure the Chrome window opened by the wizard is still open, or save the HCAD PDFs manually into the HCAD_Records folder and click Import Downloaded PDFs.')
            return

        def worker():
            try:
                result = subprocess.run(core_launch.hcad_command() + ['--file', str(address_file), '--output', str(hcad_folder), '--cdp-url', self._chrome_debug_url(), '--delay-min', '3', '--delay-max', '6'],
                                        cwd=core_launch.hcad_cwd(), creationflags=core_launch.console_flags(), text=True)
                pdfs = sorted(str(p) for p in hcad_folder.glob('*.pdf') if p.is_file())
                report = hcad_folder / 'HCAD_Download_Report.csv'
                selected = ';'.join(pdfs) if pdfs else ''
                msg = (f'HCAD download finished. {len(pdfs)} PDF(s) found in {hcad_folder}.' if pdfs else
                       f'HCAD download finished, but no PDFs were found in {hcad_folder}.\n\nIf HCAD blocked the automated download, save the HCAD PDFs manually into that folder and click Import Downloaded PDFs.')
                if report.exists():
                    msg += f'\n\nDownload report: {report}'
                if result.returncode != 0:
                    msg += f'\n\nDownloader exit code: {result.returncode}.'

                def finish():
                    if selected:
                        self.files['hcad_records'] = selected
                    self.status_var.set(msg.split(chr(10))[0])
                    messagebox.showinfo('HCAD', msg)
                self.after(0, finish)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror('HCAD', str(e)))
        threading.Thread(target=worker, daemon=True).start()

    # ---------------- pages ----------------
    def _case_info_page(self):
        self._set_active(1)
        self.clear()
        c = self._section_card(self.content, 'Case Info review', 0, 0)
        tk.Label(c, text='Case Info is read from the workbook (Case Info tab plus the live Step 2 values). Required mail-out fields: Letter_Date, Protest_Deadline, Planner_Name, Planner_Phone, Planner_Email.', bg='white', fg='#475569', font=('Segoe UI', 11), wraplength=900, justify='left').grid(row=0, column=0, sticky='w', padx=12, pady=12)
        try:
            if self.files.get('workbook'):
                new_out = maybe_case_folder(self.files.get('out', ''), self.files['workbook'])
                if new_out:
                    new_out.mkdir(parents=True, exist_ok=True)
                    self.files['out'] = str(new_out)
                    self.status_var.set(f'Output folder set to case folder: {new_out}')
            self._save_session()
        except Exception:
            pass
        snap = self._section_card(self.content, 'Case snapshot (quality check)', 1, 0)
        ci = {}
        try:
            ci = _get_case_info(load_workbook(self._wb(), data_only=True))
        except Exception as e:
            tk.Label(snap, text=f'Could not read the workbook yet: {e}', bg='white', fg='#B45309', font=('Segoe UI', 9), wraplength=860, justify='left').grid(row=0, column=0, sticky='w', padx=12, pady=10)
        if ci:
            build_snapshot_panel(snap, ci).grid(row=0, column=0, sticky='w', padx=12, pady=10)
            warnings = []
            try:
                path = _clean(ci.get('Approval_Path')) or 'Director Approval'
                n = _to_int(ci.get('Protest_Count'))
                support = _to_float(_fmt_pct_num(ci.get('Support_Percentage')) if _clean(ci.get('Support_Percentage')) else 0)
                sfr = _to_float(_fmt_pct_num(ci.get('SFR_Percentage')) if _clean(ci.get('SFR_Percentage')) else 0)
                if n > 0 and path != 'Planning Commission Approval':
                    warnings.append(f'Protest_Count is {n} but Approval_Path is "{path}". Protested applications go to the Planning Commission.')
                if 0 < support < 51 and path != 'Planning Commission Approval':
                    warnings.append(f'Support is {support:g}% (under 51%) but Approval_Path is "{path}". Applications under 51% are referred to the Planning Commission (send the notice of public hearing).')
                if 0 < sfr <= 60:
                    warnings.append(f'SFR share is {sfr:g}%: the criteria require more than 60% of the lots developed with no more than two SFR units.')
            except Exception:
                pass
            for i, w in enumerate(warnings):
                tk.Label(snap, text='⚠ ' + w, bg='white', fg='#B45309', font=('Segoe UI', 9, 'bold'), wraplength=860, justify='left').grid(row=1 + i, column=0, sticky='w', padx=12, pady=(0, 6))
        pf = self._section_card(self.content, 'Preflight checks (informational — nothing here blocks generation)', 2, 0)
        build_preflight_panel(pf, run_preflight(self.files.get('workbook'), _WIZARD_KEY, self.files.get('map'), self.files.get('out'))).grid(row=0, column=0, sticky='w', padx=12, pady=10)

    def _owners_page(self):
        self._set_active(2)
        self.clear()
        c = self._section_card(self.content, 'Owner review', 0, 0)
        tk.Label(c, text='Owner rows are read from Step 1 GIS data. Preview one letter first, then generate all mail-out letters.', bg='white', fg='#475569', font=('Segoe UI', 11), wraplength=900, justify='left').grid(row=0, column=0, sticky='w', padx=12, pady=12)
        try:
            owners = _get_owners(load_workbook(self._wb(), data_only=True))
            tk.Label(c, text=f'{len(owners)} owner letter(s) will be generated.', bg='white', fg=self.colors['text'], font=('Segoe UI', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=12, pady=(0, 6))
            for i, o in enumerate(owners[:60]):
                line = f"{i + 1}. {o['Owner_Mail']} — {o['Mail_Addr_1']}, {o['Mail_City']} {o['Mail_State']} {o['Mail_Zip']}   ({o['Property_Description']})"
                tk.Label(c, text=line, bg='white', fg='#334155', font=('Segoe UI', 9), anchor='w', wraplength=900, justify='left').grid(row=2 + i, column=0, sticky='w', padx=12, pady=1)
        except Exception as e:
            tk.Label(c, text=f'Could not read owners yet: {e}', bg='white', fg='#B45309', font=('Segoe UI', 9), wraplength=860, justify='left').grid(row=1, column=0, sticky='w', padx=12, pady=10)

    def _require_map(self):
        try:
            self._check_boundary_map()
        except Exception as e:
            messagebox.showerror('Update boundary map', str(e))
            return False
        m = self.files.get('map', '')
        if not m or not Path(m).exists():
            messagebox.showerror('Boundary map required', 'A boundary map is required to generate owner letters, because the map is embedded after each letter.\n\nSelect a boundary map on the Select Files screen, then try again.')
            return False
        return True

    def _wb(self):
        """Read the latest saved workbook; invalidate older recalculated copies."""
        if not hasattr(self, "_workbook_read_cache"):
            self._workbook_read_cache = WorkbookReadCache()
        from core.automatic_case import prepare_owner_calculation
        return self._workbook_read_cache.resolve(prepare_owner_calculation(self))

    def _save_session(self):
        try:
            core_session.save(_WIZARD_KEY, {k: v for k, v in self.files.items() if isinstance(v, str)})
        except Exception:
            pass

    def _fail(self, e):
        log_path = log_exception('MBL Planner Wizard', self.files.get('out', ''), e)
        detail = f"\n\nDetails were saved to:\n{log_path}" if log_path else ''
        messagebox.showerror('Error', str(e) + detail)
        self.status_var.set('Error: ' + str(e))

    def _notify_done(self, path, note=''):
        self._save_session()
        self.status_var.set(f'Generated: {path}' + (f'  ({note})' if note else ''))
        messagebox.showinfo('MBL Wizard', f'Generated:\n{path}' + (f'\n\n{note}' if note else ''))
        try:
            os.startfile(str(Path(path).parent))
        except Exception:
            pass

    def _letters(self, key, name, mode):
        if not self._require_map():
            return
        try:
            out = self._document_output(name)
            self._notify_done(generate_notification(self._wb(), self.files[key], self.files.get('map', ''), out, mode))
        except Exception as e:
            self._fail(e)

    def _single(self, fn, key, name, with_map=False):
        try:
            if with_map:
                self._check_boundary_map()
            out = self._document_output(name)
            if with_map:
                fn(self._wb(), self.files[key], self.files.get('map', ''), out)
            else:
                fn(self._wb(), self.files[key], out)
            self._notify_done(str(out))
        except Exception as e:
            self._fail(e)

    def _staff_report(self):
        try:
            ctx = _get_context(self._wb())
            variant = staff_report_variant(ctx)
            key = _STAFF_KEYS[variant]
            label = {'director': 'Director path (no protest, support 51% or more)', 'less51': 'support under 51% (PC hearing)', 'protest': 'protest filed (PC hearing)'}[variant]
            out = self._document_output('MBL_OneMeeting_Staff_Report.docx')
            generate_staff_report(self._wb(), self.files[key], out)
            self._notify_done(str(out), f'Staff report variant: {label}.')
        except Exception as e:
            self._fail(e)

    def _rca_packet(self):
        try:
            self._check_boundary_map()
            out = self._document_output('MBL_RCA_Packet.pdf')
            templates = {'routing': self.files['routing_template'], 'director': self.files['director_template'], 'pc': self.files['pc_template']}
            include_staff = str(self.files.get('include_staff_analysis', 'yes')).lower() in ('yes', 'true', '1')
            generate_rca_packet_pdf(self._wb(), templates, self.files.get('map', ''), self.files.get('subdivision_plat', ''), self.files.get('tax_map', ''),
                                    self.files.get('hcad_records', ''), out, include_staff_analysis=include_staff)
            self._notify_done(str(out))
        except Exception as e:
            self._fail(e)

    def _peer_review_page(self):
        self._set_active(4)
        self.clear()
        from core.peer_review_ui import reviewer_page
        reviewer_page(self, self.content)

    def _generate_page(self):
        self._set_active(3)
        self.clear()
        outer = self.content
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_columnconfigure(1, weight=1)
        tk.Label(outer, text='Step 4: Generate documents', bg=self.colors['bg'], fg=self.colors['text'], font=('Segoe UI', 16, 'bold')).grid(row=0, column=0, columnspan=2, sticky='w', padx=8, pady=(0, 6))
        btn = dict(bg='#E7E5E4', font=('Segoe UI', 10, 'bold'), padx=18, pady=12, relief='solid', bd=1)
        c = self._section_card(outer, 'Owner Notification Mail-Out', 1, 0, 2)
        tk.Label(c, text='One preview letter or one combined Word document with every owner letter; the boundary map follows each letter.', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=900, justify='left').grid(row=0, column=0, columnspan=2, sticky='w', padx=12, pady=(4, 10))
        tk.Button(c, text='Preview One Notification Letter', command=lambda: self._letters('template', 'MBL_Notification_Preview.docx', 'preview'), **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        tk.Button(c, text='Generate All Notification Letters', command=lambda: self._letters('template', 'MBL_All_Owner_Notification_Letters.docx', 'all'), **btn).grid(row=1, column=1, sticky='w', padx=12, pady=8)
        tk.Label(c, text='Support under 51%: the notice of public hearing (Hearing_Letter_Date, PC_Hearing_Date). Protest filed: the protest letter (Protest_Letter_Date, PC_Hearing_Date). Hearing moved: the notice of PC rescheduling (Reschedule_Letter_Date, new PC_Hearing_Date).', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=900, justify='left').grid(row=2, column=0, columnspan=2, sticky='w', padx=12, pady=(6, 6))
        tk.Button(c, text='Preview Public Hearing Notice (under 51%)', command=lambda: self._letters('hearing_template', 'MBL_Public_Hearing_Notice_Preview.docx', 'preview'), **btn).grid(row=3, column=0, sticky='w', padx=12, pady=8)
        tk.Button(c, text='Generate All Public Hearing Notices', command=lambda: self._letters('hearing_template', 'MBL_All_Owner_Public_Hearing_Notices.docx', 'all'), **btn).grid(row=3, column=1, sticky='w', padx=12, pady=8)
        tk.Button(c, text='Preview Protest Letter', command=lambda: self._letters('protest_template', 'MBL_Protest_Letter_Preview.docx', 'preview'), **btn).grid(row=4, column=0, sticky='w', padx=12, pady=8)
        tk.Button(c, text='Generate All Protest Letters', command=lambda: self._letters('protest_template', 'MBL_All_Owner_Protest_Letters.docx', 'all'), **btn).grid(row=4, column=1, sticky='w', padx=12, pady=8)
        tk.Button(c, text='Preview Rescheduling Notice', command=lambda: self._letters('reschedule_template', 'MBL_Rescheduling_Notice_Preview.docx', 'preview'), **btn).grid(row=5, column=0, sticky='w', padx=12, pady=8)
        tk.Button(c, text='Generate All Rescheduling Notices', command=lambda: self._letters('reschedule_template', 'MBL_All_Owner_Rescheduling_Notices.docx', 'all'), **btn).grid(row=5, column=1, sticky='w', padx=12, pady=8)
        app = self._section_card(outer, 'Applicant Letter', 2, 0)
        tk.Label(app, text='Complete / sign-posting letter to the applicant (Letter_Date is the completion date).', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=430, justify='left').grid(row=0, column=0, columnspan=2, sticky='w', padx=12, pady=(4, 10))
        tk.Button(app, text='Generate Applicant Sign Letter', command=lambda: self._single(generate_applicant_letter, 'applicant_template', 'MBL_Applicant_Sign_Letter.docx'), **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        appr = self._section_card(outer, 'Approvals', 2, 1)
        tk.Label(appr, text='Director Approval (support 51% or more, no protest) or Planning Commission Approval (under 51% or protest filed).', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=430, justify='left').grid(row=0, column=0, columnspan=2, sticky='w', padx=12, pady=(4, 10))
        tk.Button(appr, text='Generate Director Approval', command=lambda: self._single(generate_director_approval, 'director_template', 'MBL_Director_Approval.docx'), **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        tk.Button(appr, text='Generate PC Approval', command=lambda: self._single(generate_pc_approval, 'pc_template', 'MBL_PC_Approval.docx'), **btn).grid(row=1, column=1, sticky='w', padx=12, pady=8)
        forms = self._section_card(outer, 'RCA Packet Forms', 3, 0)
        tk.Label(forms, text='Routing form (checks Building Line and SMBL Block Application, fills support, SFR and lot counts) and the OneMeeting staff report. The staff report picks the Director, under-51% or protest version from Case Info.', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=430, justify='left').grid(row=0, column=0, columnspan=2, sticky='w', padx=12, pady=(4, 10))
        tk.Button(forms, text='Generate Routing Form', command=lambda: self._single(generate_routing_form, 'routing_template', 'MBL_Routing_Form.docx'), **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        tk.Button(forms, text='Generate OneMeeting Staff Report', command=self._staff_report, **btn).grid(row=1, column=1, sticky='w', padx=12, pady=8)
        ordc = self._section_card(outer, 'Draft Ordinance', 3, 1)
        tk.Label(ordc, text='Planner draft ordinance with Exhibit A property descriptions, Exhibit B map and Exhibit C property list (standalone Word document).', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=430, justify='left').grid(row=0, column=0, sticky='w', padx=12, pady=(4, 10))
        tk.Button(ordc, text='Generate Draft Ordinance', command=lambda: self._single(generate_draft_ordinance, 'ordinance_template', 'MBL_Draft_Ordinance.docx', with_map=True), **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        packet = self._section_card(outer, 'RCA Packet PDF Builder', 4, 0, 2)
        tk.Label(packet, text='Routing form, Director or PC approval, boundary map, Step 2 staff analysis, plat (when selected), tax map and HCAD records (when selected).', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=900, justify='left').grid(row=0, column=0, sticky='w', padx=12, pady=(4, 10))
        tk.Button(packet, text='Generate MBL RCA Packet PDF', command=self._rca_packet, bg='#DCFCE7', activebackground='#86EFAC', font=('Segoe UI', 10, 'bold'), padx=18, pady=12, relief='solid', bd=1).grid(row=1, column=0, sticky='w', padx=12, pady=8)

        pc = self._section_card(outer, 'Planning Commission Report and Presentation', 5, 0, 2)
        tk.Label(pc, text='Prepare the PC report and speech from the calculation sheet. The PC map uses the standard colors, symbols and editing controls, with building lines in feet.', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=900, justify='left').grid(row=0, column=0, columnspan=3, sticky='w', padx=12, pady=(4, 10))
        tk.Button(pc, text='Generate PC Map', command=lambda: self._open_pc_draft(start_map=True), **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        tk.Button(pc, text='Generate PC Draft Report', command=self._open_pc_draft, **btn).grid(row=1, column=1, sticky='w', padx=12, pady=8)
        tk.Button(pc, text='PC Speech', command=self._open_pc_speech, **btn).grid(row=1, column=2, sticky='w', padx=12, pady=8)

    def _open_pc_draft(self, start_map=False):
        try:
            from wizards.mls.pc_ui import open_pc_workspace
            open_pc_workspace(self, start_map)
        except Exception as error:
            messagebox.showerror('Planning Commission', str(error), parent=self)

    def _open_pc_speech(self):
        try:
            from wizards.mls.pc_speech_ui import open_pc_speech
            open_pc_speech(self)
        except Exception as error:
            messagebox.showerror('PC Speech', str(error), parent=self)


if __name__ == '__main__':
    print('This wizard now runs inside the RCA Wizard Suite.')
    print('Start it with Run Wizard.bat, or: python app.py --wizard mbl')
