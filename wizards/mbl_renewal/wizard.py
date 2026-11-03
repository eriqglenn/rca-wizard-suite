"""MBL Renewal Planner Wizard (v2.0, 2026 City templates).

Fills the official City of Houston Special Minimum Building Line Block
*renewal* templates from the renewal calculation workbook (Step 1 GIS data,
Step 2 calculation sheet, Case Info tab) and assembles the RCA packet.

Documents generated (templates in templates/mbl_renewal, filled through the
shared wording-based engine in core/docx_merge.py):

    Owner notification letters     mail merge, map after each letter
    Protest letter to owners       public-hearing notice when a protest is filed
    Notice of PC rescheduling      new hearing date to all owners
    Application acceptance letter  to the applicant
    Applicant sign letter          completion / sign-posting letter
    Director approval              2026 renewal criteria form
    Planning Commission approval   renewal criteria form (unchanged City form)
    Routing form                   2026 coversheet routing form (content controls)
    OneMeeting staff report        Director-path NOVUS cover letter
    Draft renewal ordinance        standalone Word document
    RCA packet PDF                 routing, approval, map, tax map, original ordinance
"""
import tkinter as tk
import datetime
import re
from tkinter import filedialog, messagebox
from pathlib import Path
import tempfile, os, shutil, subprocess, sys, zipfile, copy, threading, time, urllib.request
from copy import deepcopy
from collections import Counter
from openpyxl import load_workbook
from core.routing_form import prepare_routing_form, routing_pdf_one_page
from docx import Document
from docx.oxml.ns import qn
from docx.shared import RGBColor

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
from core.sheet_layout import apply_staff_analysis_page_setup
from core.docx_merge import (
    _RunCursor, _append_letter_as_own_section, _cap, _clean, _clear_all_red_run_colors,
    _compact_notification_first_page, _fill_content_controls, _fill_phone, _fix_mailto_links,
    _fmt_int, _fmt_pct, _fmt_pct_num, _insert_map, _iter_all_paragraphs, _map_to_png, _norm,
    _num_words, _parse_date, _replace_in_runs, _replace_runs_contextual, _replace_text_nodes,
    _restart_list_numbering, _set_exact_para_preserve_first_run, _set_paragraph_checkboxes,
    _set_run_text_preserve_xml, _blacken_run, _to_int, _count_phrase, _EMAIL_PLACEHOLDERS,
    _EXCEL_ERRORS, _force_paragraph_color_black,
)


def _fill_runs(p, exact_map=None, contains_map=None):
    """Run-level fill that also clears the template's red placeholder colour
    from the rest of the paragraph (the City letters colour the fixed words
    around a placeholder, e.g. 'XX feet', 'DATE, at 5:00 p.m.')."""
    if _replace_runs_contextual(p, exact_map=exact_map, contains_map=contains_map):
        _force_paragraph_color_black(p)
        return True
    return False
# Ordinance paragraph/exhibit helpers are shared with the MLS Renewal wizard.
from wizards.mls_renewal.wizard import (
    _replace_first_content_control_text, _set_ordinance_para, _set_whereas,
    _populate_exhibit_a, _populate_exhibit_b, _populate_exhibit_c, _remove_red_from_doc,
    _normalize_draft_ordinance_fonts, _strip_all_content_controls,
    _remove_paragraph_section_breaks, _get_property_rows, _step1_reader,
    _pdf_page_is_blank, _merge_pdfs, _ensure_pdf, _convert_office_to_pdf,
)

WIZARD_VERSION = 'v2.0'
STEP2_SHEET = 'Step 2 - Main Calculation Sheet'
STEP1_SHEET = 'Step 1 - Input Properties Info'


# =============================================================================
# Workbook readers
# =============================================================================
def _step2_layout(ws):
    """'new' for the 2026 City sheet (D2 = building line to renew, column F
    flags), 'old' for the pre-2026 sheet (E10 most frequent line, H9 share)."""
    c2 = _clean(ws['C2'].value).lower()
    if 'building line to renew' in c2 or _to_int(ws['D2'].value) > 0:
        return 'new'
    if 'most frequent' in _clean(ws['D10'].value).lower() or _to_int(ws['E10'].value) > 0:
        return 'old'
    return 'new'


def _get_step2(wb):
    ws = wb[STEP2_SHEET] if STEP2_SHEET in wb.sheetnames else wb[wb.sheetnames[1]]
    layout = _step2_layout(ws)
    rows = []
    for r in range(15, min(ws.max_row, 140) + 1):
        desc = _clean(ws.cell(r, 1).value)
        cur = _to_int(ws.cell(r, 4).value)
        if not desc and cur <= 0:
            continue
        rows.append({
            'desc': desc,
            'land_use': _clean(ws.cell(r, 2).value).upper(),
            'orig': _to_int(ws.cell(r, 3).value),
            'cur': cur,
            'lot': _to_int(ws.cell(r, 5).value),
            'flag': _clean(ws.cell(r, 6).value).upper() if layout == 'new' else '',
        })
    if not rows:
        # Stale formula cache: gate rows and land use from the raw Step 1 columns
        # the Step 2 mirror formulas read; building lines still come from column D.
        try:
            s1 = wb[STEP1_SHEET]
            headers = [_clean(s1.cell(1, c).value) for c in range(1, s1.max_column + 1)]
            idx = {h: c + 1 for c, h in enumerate(headers)}
            for i in range(0, 125):
                r1, r2 = 2 + i, 15 + i
                num = _clean(s1.cell(r1, idx.get('Str_num', 5)).value)
                legal = _clean(s1.cell(r1, idx.get('Legal_Dscr_1', 30)).value)
                if not num and not legal:
                    continue
                grp = _clean(s1.cell(r1, idx.get('group_dscr', 35)).value).lower()
                lu = 'SFR' if 'single-family' in grp else ('VAC' if 'undeveloped' in grp else (
                    'COM' if 'commercial' in grp else ('MF' if 'multi-family' in grp else 'EXC')))
                rows.append({'desc': f"{num} ({legal})", 'land_use': lu, 'orig': _to_int(ws.cell(r2, 3).value),
                             'cur': _to_int(ws.cell(r2, 4).value), 'lot': _to_int(s1.cell(r1, idx.get('Total_Land_Area', 41)).value), 'flag': ''})
        except Exception:
            pass

    if layout == 'old':
        min_bl = _to_int(ws['E10'].value)
        sfr_pct_cell = ws['H9'].value
        total = _to_int(ws['F9'].value) or _to_int(ws['A13'].value) or len(rows)
        sfr_count = _to_int(ws['D9'].value)
        qualifies = _clean(ws['E2'].value)
    else:
        min_bl = _to_int(ws['D2'].value)
        sfr_pct_cell = ws['D9'].value
        total = _to_int(ws['A13'].value) or len(rows)
        sfr_count = _to_int(ws['A9'].value)
        qualifies = _clean(ws['D10'].value)
    sfr_rows = [x for x in rows if x['land_use'] == 'SFR']
    if not sfr_count:
        sfr_count = len(sfr_rows)
    if min_bl <= 0:
        bls = [x['cur'] for x in sfr_rows if x['cur'] > 0]
        min_bl = Counter(bls).most_common(1)[0][0] if bls else 0
    try:
        sfr_pct = float(sfr_pct_cell)
    except Exception:
        sfr_pct = (sfr_count / total) if total else 0
    meeting = sum(1 for x in rows if x['flag'] == 'YES')
    if not meeting and rows and not any(x['flag'] for x in rows):
        meeting = sum(1 for x in rows if x['cur'] > 0 and (x['cur'] >= min_bl or (x['orig'] and x['cur'] == x['orig'])))
    compliant = sum(1 for x in sfr_rows if min_bl and x['cur'] == min_bl)
    vary = sum(1 for x in sfr_rows if min_bl and x['cur'] > 0 and abs(x['cur'] - min_bl) > 5)
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
        'Minimum_Building_Line_Text': _fmt_int(min_bl) if min_bl else '',
        'SFR_Count': sfr_count,
        'SFR_Lot_Count': sfr_count,
        'SFR_Percentage': _fmt_pct(sfr_pct),
        'Total_Lot_Count': total,
        'Lots_Meeting_Minimum': meeting,
        'Lots_Meeting_Percentage': _fmt_pct((meeting / total) if total else 0),
        'Compliant_BL_Count': compliant,
        'Variance_Count': vary,
        'Variance_Percentage': _fmt_pct((vary / sfr_count) if sfr_count else 0),
        'Qualifies': qualifies,
        'Step2_Layout': layout,
        'Property_Rows': rows,
    }
    if case['Application_Number'] in ('0', '0.0'):
        case['Application_Number'] = ''
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
            'Application_Type': 'Special Minimum Building Line Block Renewal Application',
            'Date_Received': s2.get('Date_Received', ''),
            'Date_Complete': s2.get('Date_Complete', ''),
            'Street_Name': s2.get('Street_Name', ''),
            'Block_Description': s2.get('Block_Description', ''),
            'Side_Of_Street': s2.get('Side_Text', ''),
            'Cross_Streets': (s2.get('Area_Line_2', '') or '').replace('between ', '').strip(),
            'Minimum_Building_Line': s2.get('Minimum_Building_Line_Text', ''),
            'SFR_Count': s2.get('SFR_Count', ''),
            'SFR_Percentage': s2.get('SFR_Percentage', ''),
            'Total_Lot_Count': s2.get('Total_Lot_Count', ''),
            'Lots_Meeting_Minimum': s2.get('Lots_Meeting_Minimum', ''),
            'Lots_Meeting_Percentage': s2.get('Lots_Meeting_Percentage', ''),
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
            'Total_Land_Area': _fmt_int(gv(r, 'Total_Land_Area')), 'group_dscr': gv(r, 'group_dscr'),
            'Property_Description': gv(r, 'Property Description'),
        })
    from core.lot_divisions import collapse_lot_notifications
    return collapse_lot_notifications(wb,owners)


def _get_renewal_context(workbook_path):
    wb = load_workbook(workbook_path, data_only=True)
    ci = _get_case_info(wb)
    s = _get_step2(wb)
    app = str(s.get('Application_Number') or _clean(ci.get('Application_Number')) or '').strip()
    if app.upper().endswith('REN'):
        app_ren, app_prefix = app, app[:-3]
    else:
        app_ren, app_prefix = (f'{app}REN' if app else 'XXREN'), (app or 'XX')
    block = s.get('Block_Description', '')
    side_label = s.get('Side_Label', '')
    c1, c2 = s.get('Cross_Street_1', ''), s.get('Cross_Street_2', '')
    between = f"between {c1} and {c2}" if (c1 or c2) else ''
    area = ', '.join(x for x in [block, side_label, between] if x)
    caption = (f"{side_label} of the {block}" if side_label else block)
    if between:
        caption += f", {between}"
    min_bl = s.get('Minimum_Building_Line') or 0
    total = _to_int(s.get('Total_Lot_Count'))
    sfr = _to_int(s.get('SFR_Count'))
    meeting = _to_int(s.get('Lots_Meeting_Minimum'))
    compliant = _to_int(s.get('Compliant_BL_Count'))
    vary = _to_int(s.get('Variance_Count'))
    protest_n = _to_int(ci.get('Protest_Count'), 0)
    if protest_n <= 0:
        protest_sentence = 'No protest was received.'
    elif protest_n == 1:
        protest_sentence = 'One (1) protest was filed.'
    else:
        protest_sentence = f'{_count_phrase(protest_n)} protests were filed.'
    eff = _clean(ci.get('Original_Effective_Date'))
    eff_date = _parse_date(eff)
    hearing = _clean(ci.get('PC_Hearing_Date'))
    hearing_date = _parse_date(hearing)
    context = {**{k: _clean(v) for k, v in ci.items()}, **s}
    context.update({
        'Application_Number': app or 'XX',
        'Application_Renewal_Number': app_ren,
        'Application_Prefix': app_prefix,
        'Area_Description': area,
        'Area_Description_Upper': area.upper(),
        'Ordinance_Caption': caption,
        'Ordinance_Caption_Upper': caption.upper(),
        'Between_Streets': between,
        'Applicant_Name': _clean(ci.get('Applicant_Name')) or 'Applicant Name',
        'Applicant_Address_Block': '\n'.join(x for x in [_clean(ci.get('Applicant_Address_1')), _clean(ci.get('Applicant_Address_2'))] if x) or 'Address',
        'Applicant_City': _clean(ci.get('Applicant_City')) or 'Houston',
        'Applicant_State': _clean(ci.get('Applicant_State')) or 'TX',
        'Applicant_Zip': _clean(ci.get('Applicant_Zip')) or '',
        'Total_Lot_Count': str(total) if total else 'XX',
        'Total_Lot_Count_Words': _num_words(total) if total else 'XX',
        'SFR_Lot_Count': str(sfr) if sfr else 'X',
        'SFR_Lot_Count_Words': _num_words(sfr) if sfr else 'X',
        'SFR_Percentage_Number': _fmt_pct_num(s.get('SFR_Percentage')) or 'X',
        'Minimum_Building_Line_Number': str(min_bl) if min_bl else 'XX',
        'Minimum_Building_Line_Text': _fmt_int(min_bl) if min_bl else 'XX',
        'Minimum_Building_Line_Words': _num_words(min_bl) if min_bl else 'X',
        'Qualifying_Lot_Count': str(meeting) if meeting else 'X',
        'Compliant_BL_Phrase': _count_phrase(compliant) if compliant else 'X (X)',
        'Variance_Phrase': _count_phrase(vary) if vary else 'Zero (0)',
        'Variance_Percentage_Number': _fmt_pct_num(s.get('Variance_Percentage')) or '0',
        'Protest_Count': protest_n,
        'Protest_Sentence': protest_sentence,
        'Original_Ordinance_No': _clean(ci.get('Original_Ordinance_No')) or 'XXXX-XX',
        'Original_Effective_Date': eff or 'Month XX, XXXX',
        'Original_Ordinance_Year': str(eff_date.year) if eff_date else (_clean(ci.get('Original_Ordinance_No')).split('-')[0] or 'XXXX'),
        'Original_Expiration_Date': _clean(ci.get('Original_Expiration_Date')) or 'Month XX, XXXX',
        'Application_Filed_Date': _clean(ci.get('Application_Filed_Date')) or _clean(s.get('Date_Received')) or 'Month XX, XXXX',
        'Application_Accepted_Date': _clean(ci.get('Application_Accepted_Date')) or _clean(s.get('Date_Received')) or 'Date',
        'Application_Complete_Date': _clean(ci.get('Application_Complete_Date')) or _clean(s.get('Date_Complete')) or 'Month XX, XXXX',
        'Council_Deadline': _clean(ci.get('Council_Deadline')) or '',
        'Council_District': _clean(ci.get('Council_District')) or '',
        'Council_Member': _clean(ci.get('Council_Member')).upper() or 'CM NAME',
        'Approval_Path': _clean(ci.get('Approval_Path')) or 'Director Approval',
        'Planner_Name': _clean(ci.get('Planner_Name')) or 'Planner Name',
        'Planner_Position': _clean(ci.get('Planner_Position')) or 'Planner',
        'Planner_Phone': _clean(ci.get('Planner_Phone')) or '832-393-XXXX',
        'Planner_Email': _clean(ci.get('Planner_Email')) or 'Planner.Name@houstontx.gov',
        'Letter_Date': _clean(ci.get('Letter_Date')) or 'Date',
        'Protest_Letter_Date': _clean(ci.get('Protest_Letter_Date')) or _clean(ci.get('Letter_Date')) or 'Date',
        'Reschedule_Letter_Date': _clean(ci.get('Reschedule_Letter_Date')) or _clean(ci.get('Letter_Date')) or 'Date',
        'Protest_Deadline': _clean(ci.get('Protest_Deadline')) or 'DATE',
        'PC_Hearing_Date': hearing or '[Date]',
        'PC_Hearing_Weekday': (hearing_date.strftime('%A') + ',') if hearing_date else 'Thursday,',
        'Support_Percentage': _fmt_pct_num(ci.get('Support_Percentage')) if _clean(ci.get('Support_Percentage')) else 'XX',
        'Director_Name': _clean(ci.get('Director_Name')) or 'Vonn Tran',
        'Applicant_Property_Address': _clean(ci.get('Applicant_Property_Address')) or 'XX Street',
        'Subdivision_Name': _clean(ci.get('Subdivision_Name')) or '(Subdivision Name)',
        'Sign_Count_Number': _clean(ci.get('Sign_Count_Number')) or '2',
        'Sign_Count_Words': _clean(ci.get('Sign_Count_Words')) or 'two',
        'Sign_Placement': _clean(ci.get('Sign_Placement')) or 'one sign at each corner',
        'Properties': _get_property_rows(wb),
    })
    context['Blockfaces_Sentence'] = (
        f"The block is comprised of {context['Total_Lot_Count_Words']} ({context['Total_Lot_Count']}) lots. "
        f"{context['Compliant_BL_Phrase']} of the {context['SFR_Lot_Count_Words']} ({context['SFR_Lot_Count']}) SFR lots have a constructed building line of "
        f"{context['Minimum_Building_Line_Text']} feet. The lots that are developed for single-family residential use "
        f"comply with the provisions of the criteria below.")
    context['Variance_Sentence'] = (
        f"The application meets this standard. {context['Variance_Phrase']} of {context['SFR_Lot_Count_Words']} ({context['SFR_Lot_Count']}) SFR lots have a building line that varies by more than "
        f"five (5) feet from the most frequently constructed building line, representing "
        f"{context['Variance_Percentage_Number']} percent of the total SFR lots.")
    return context


# =============================================================================
# Owner letters (notification / protest / PC rescheduling notice)
# =============================================================================
def _fill_owner_letter(doc, ctx, owner):
    full = _norm(' '.join(p.text for p in doc.paragraphs))
    if 'received a protest' in full:
        kind, letter_date = 'protest', ctx['Protest_Letter_Date']
    elif 'rescheduled' in full:
        kind, letter_date = 'reschedule', ctx['Reschedule_Letter_Date']
    else:
        kind, letter_date = 'notice', ctx['Letter_Date']
    min_bl = ctx['Minimum_Building_Line_Text']
    addr = owner.get('Mail_Addr_1', '')
    if owner.get('Mail_Addr_2') and owner.get('Mail_Addr_2') not in ('<Null>', '0'):
        addr += '\n' + owner['Mail_Addr_2']
    common = {
        '«Owner_Mail»': owner.get('Owner_Mail', ''),
        '«Mail_Addr_»': addr,
        '«Mail_City»': owner.get('Mail_City', ''),
        '«Mail_State»': owner.get('Mail_State', ''),
        '«Mail_Zip»': owner.get('Mail_Zip', ''),
        'XX block of XX Street, X and X sides,': ctx['Area_Line_1'] or 'XX block of XX Street, X and X sides,',
        'between X and X': ctx['Area_Line_2'] or 'between X and X',
    }
    for p in doc.paragraphs:
        t = _norm(p.text)
        ts = t.strip()
        if ts == 'DATE':
            _set_exact_para_preserve_first_run(p, letter_date)
            continue
        if '2:30' in t and 'DATE' in t:
            _fill_runs(p, exact_map={'DATE': ctx['PC_Hearing_Date']})
        elif 'protest' in t and 'DATE' in t:
            _fill_runs(p, exact_map={'DATE': ctx['Protest_Deadline']}, contains_map={'DATE': ctx['Protest_Deadline']})
        if 'public hearing is scheduled' in t:
            _fill_runs(p, exact_map={'Thursday,': ctx['PC_Hearing_Weekday']},
                                     contains_map={'[Date]': ctx['PC_Hearing_Date']})
        if '% of property owners' in t:
            _fill_runs(p, exact_map={'XX': ctx['Support_Percentage']})
        elif 'feet' in t:
            _fill_runs(p, exact_map={'XX': min_bl})
        if 'Planner' in t or '393' in t:
            _fill_runs(p, exact_map={'Planner Name': ctx['Planner_Name']},
                                     contains_map={'Planner Name': ctx['Planner_Name'], 'contact Planner,': f"contact {ctx['Planner_Name']},",
                                                   'or call the ': 'or call '})
            _fill_phone(p, ctx['Planner_Phone'])
        _replace_in_runs(p, common)
    _replace_text_nodes(doc, {k: ctx['Planner_Email'] for k in _EMAIL_PLACEHOLDERS})
    _fix_mailto_links(doc, ctx['Planner_Email'])


def generate_notification(workbook_path, template_path, map_path, output_path, mode='preview'):
    ctx = _get_renewal_context(workbook_path)
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
# Applicant letters (acceptance / complete-and-sign)
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
    ctx = _get_renewal_context(workbook_path)
    doc = Document(template_path)
    common = _applicant_common(ctx)
    for p in doc.paragraphs:
        t = _norm(p.text)
        ts = t.strip()
        if ts == 'DATE':
            _set_exact_para_preserve_first_run(p, ctx['Letter_Date'])
            continue
        if 'Minimum Building Line is' in t:
            _fill_runs(p, exact_map={
                'XX': ctx['Minimum_Building_Line_Text'],
                'two (2': f"{ctx['Sign_Count_Words']} ({ctx['Sign_Count_Number']}",
                'one': ctx['Sign_Placement']}, contains_map={' sign at each corner.': '.'})
        elif 'feel free to contact' in t:
            _fill_runs(p, exact_map={'Planner': ctx['Planner_Name'], 'xxxx': '', 'XXXX': ''},
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


def generate_acceptance_letter(workbook_path, template_path, output_path):
    ctx = _get_renewal_context(workbook_path)
    doc = Document(template_path)
    common = _applicant_common(ctx)
    accepted = ctx['Application_Accepted_Date']
    side_label = ctx.get('Side_Label', '')
    block = ctx.get('Block_Description', '')
    for p in doc.paragraphs:
        t = _norm(p.text)
        ts = t.strip()
        if ts == 'DATE':
            _set_exact_para_preserve_first_run(p, accepted)
            continue
        if ts.startswith('Dear'):
            _fill_runs(p, exact_map={'Property Owners Name': ctx['Applicant_Name']})
        elif 'accepted your application' in t:
            c = _RunCursor(p)
            c.set('Date', accepted)
            c.span('XX', '(X side)', f"{block} ({side_label})" if side_label else block)
            c.set(' Between ', ' between ')
            c.set('XX', ctx.get('Cross_Street_1', '') or 'XX')
            c.set('XX', ctx.get('Cross_Street_2', '') or 'XX')
            c.set('XXX', ctx['Application_Renewal_Number'])
            c.set('Planner’s Name', ctx['Planner_Name'])
        elif 'feel free to contact' in t:
            _fill_runs(p, exact_map={'Planner': ctx['Planner_Name'], 'XXXX': '',
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
def _replace_line_in_paragraph(p, marker, new_line):
    """Replace the single <w:br/>-separated line containing `marker` inside a
    criteria cell (criterion statement, blank line, finding) and keep the rest."""
    joined = p.text
    if marker not in joined:
        return False
    lines = joined.split('\n')
    for i, line in enumerate(lines):
        if marker in line:
            lines[i] = new_line
            break
    first = p.runs[0] if p.runs else None
    rpr = copy.deepcopy(first._r.find(qn('w:rPr'))) if first is not None and first._r.find(qn('w:rPr')) is not None else None
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    run = p.add_run()
    if rpr is not None:
        run._r.insert(0, rpr)
    for i, line in enumerate(lines):
        if i > 0:
            run.add_break()
        if line:
            run.add_text(line)
    _blacken_run(run)
    return True


def generate_director_approval(workbook_path, template_path, output_path):
    ctx = _get_renewal_context(workbook_path)
    from core.city_templates import render_updated
    if render_updated(template_path, output_path, ctx, workbook_path, 'mbl_renewal', 'director'): return 1
    doc = Document(template_path)
    from core.mbl_templates_2026 import fill_mbl_template
    if fill_mbl_template(doc, ctx, renewal=True, routing=False):
        _clear_all_red_run_colors(doc)
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out))
        unlock_docx(out)
        return 1
    for p in _iter_all_paragraphs(doc):
        t = _norm(p.text)
        ts = t.strip()
        if 'Block Renewal No.' in t:
            _replace_runs_contextual(p, exact_map={'XX': ctx['Application_Prefix'], 'XXREN': ctx['Application_Renewal_Number']})
        elif 'The application is for the' in t:
            c = _RunCursor(p)
            if not c.span('XX', 'Street', ctx['Area_Description'] or 'application area'):
                _replace_line_in_paragraph(p, 'The application is for the', f"The application is for the {ctx['Area_Description']}.")
        elif 'went into effect on' in t:
            c = _RunCursor(p)
            c.span('200', '63', ctx['Original_Ordinance_No'])
            c.set('January 31, 2006', ctx['Original_Effective_Date'])
            _replace_runs_contextual(p, contains_map={'2006-63': ctx['Original_Ordinance_No'], 'January 31, 2006': ctx['Original_Effective_Date']})
        elif 'minimum building line of' in t:
            _replace_runs_contextual(p, exact_map={'XX': ctx['Minimum_Building_Line_Text']})
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


def _pc_approval_fields(ctx):
    """2026 approval slots; the noncorner count is a planner-entered fact."""
    raw = _clean(ctx.get('PC_Noncorner_Lot_Count'))
    count = None
    try:
        numeric = float(raw)
        if numeric.is_integer() and 0 <= numeric <= _to_int(ctx.get('Total_Lot_Count')):
            count = int(numeric)
    except (TypeError, ValueError, OverflowError):
        pass
    return {
        '{{MBL_REN_NUMBER}}': ctx['Application_Renewal_Number'],
        '{{MBL_REN_AREA}}': ctx['Area_Description'] or '[ENTER APPLICATION AREA]',
        '{{MBL_REN_SFR_PERCENT}}': ctx['SFR_Percentage_Number'],
        '{{MBL_REN_PROTEST}}': ctx['Protest_Sentence'],
        '{{MBL_REN_TOTAL_LOTS}}': f"{ctx['Total_Lot_Count_Words']} ({ctx['Total_Lot_Count']})",
        '{{MBL_REN_NONCORNER_LOTS}}': _count_phrase(count) if count is not None else '[ENTER COUNT]',
        '{{MBL_REN_BUILDING_LINE}}': ctx['Minimum_Building_Line_Text'],
        '{{MBL_REN_PC_DATE}}': _clean(ctx.get('PC_Approval_Date')) or 'Date',
    }


def generate_pc_approval(workbook_path, template_path, output_path):
    ctx = _get_renewal_context(workbook_path)
    from core.city_templates import render_updated
    if render_updated(template_path, output_path, ctx, workbook_path, 'mbl_renewal', 'pc'): return 1
    doc = Document(template_path)
    if any('{{MBL_REN_' in (t.text or '') for t in doc.element.xpath('.//w:t')):
        # Every slot is one text node so all City run formatting stays intact.
        _replace_text_nodes(doc, _pc_approval_fields(ctx))
        _clear_all_red_run_colors(doc)
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out))
        unlock_docx(out)
        return 1
    # Continue supporting a planner-selected older City template.
    for p in _iter_all_paragraphs(doc):
        t = _norm(p.text)
        ts = t.strip()
        if 'Block Renewal No.' in t:
            _replace_in_runs(p, {'XXREN': ctx['Application_Renewal_Number']})
        elif 'The application is for the' in t:
            _replace_line_in_paragraph(p, 'The application is for the', f"The application is for the {ctx['Area_Description']}.")
        elif '% of the lots are developed' in t:
            _replace_line_in_paragraph(p, '% of the lots are developed', f"{ctx['SFR_Percentage_Number']}% of the lots are developed with single-family residential units.")
        elif 'protest was filed' in t:
            _replace_line_in_paragraph(p, 'protest was filed', f"Evidence of support is not required per the ordinance. All property owners were notified that a protest of the application must be filed within thirty (30) days. {ctx['Protest_Sentence']}")
        elif 'The block is comprised of' in t:
            _replace_line_in_paragraph(p, 'The block is comprised of', ctx['Blockfaces_Sentence'])
        elif 'The application meets this standard' in t:
            _replace_line_in_paragraph(p, 'The application meets this standard', ctx['Variance_Sentence'])
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
def generate_routing_form(workbook_path, template_path, output_path):
    ctx = _get_renewal_context(workbook_path)
    from core.city_templates import render_updated
    if render_updated(template_path, output_path, ctx, workbook_path, 'mbl_renewal', 'routing'): return 1
    doc = Document(template_path)
    from core.mbl_templates_2026 import fill_mbl_template
    if fill_mbl_template(doc, ctx, renewal=True, routing=True):
        _clear_all_red_run_colors(doc)
        prepare_routing_form(doc)
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out))
        unlock_docx(out)
        return 1
    is_pc = ctx['Approval_Path'] == 'Planning Commission Approval'
    location = f"The {ctx['Area_Description']}" if ctx.get('Block_Description') else 'The XX block of X Street, X and X sides, between XX and XX'
    rules = [
        ('Special Minimum Building Line Block Application No.XX',
         f"Special Minimum Building Line Block Renewal Application No. {ctx['Application_Renewal_Number']}", None),
        ('Special Minimum Building Line/ Block Application No.XX',
         f"Special Minimum Building Line Block Renewal Application No. {ctx['Application_Renewal_Number']}", None),
        ('The XX block of X Street, X and X sides, between XX and XX', location, None),
        ('Planner Name, Phone Number', f"{ctx['Planner_Name']}, {ctx['Planner_Phone']}", None),
        ('DATE', ctx['Council_Deadline'] or 'DATE', 'no later than'),
        ('Previous ordinance: XXXX', f"Previous ordinance: {ctx['Original_Ordinance_No']}", None),
        ('XXX', ctx['Application_Renewal_Number'], 'BUILDING LINE'),
        (': XX', f": {ctx['Minimum_Building_Line_Text']}", 'Line Established'),
        ('XX', ctx['SFR_Percentage_Number'], 'SFR:'),
        ('XX', 'N/A', 'Petition'),
        ('XX / XX', f"{ctx['Total_Lot_Count']} / {ctx['Qualifying_Lot_Count']}", None),
    ]
    _fill_content_controls(doc, rules)
    for i, p in enumerate(doc.paragraphs):
        t = _norm(p.text)
        if t.startswith('The topic of Request') and i + 1 < len(doc.paragraphs):
            loc = doc.paragraphs[i + 1]
            if loc.text.strip().startswith('('):
                ind = loc._p.get_or_add_pPr().find(qn('w:ind'))
                if ind is not None:
                    left = int(ind.get(qn('w:left'), '0') or 0) + int(ind.get(qn('w:firstLine'), '0') or 0)
                    ind.set(qn('w:left'), str(left))
                    ind.set(qn('w:firstLine'), '0')
                    if qn('w:hanging') in ind.attrib:
                        del ind.attrib[qn('w:hanging')]
        elif 'BUILDING LINE' in t and 'LOT SIZE' in t:
            _set_paragraph_checkboxes(p, [True, False])
        elif t.startswith('Petition Signed By'):
            # renewals are not petition-driven: drop the literal '%' after N/A
            _replace_runs_contextual(p, exact_map={'%': ''})
        elif 'SMLS Area Application' in t:
            _set_paragraph_checkboxes(p, [False, False, False])
        elif 'SMLS Block Renewal' in t and 'SMBL' not in t:
            _set_paragraph_checkboxes(p, [False])
        elif 'SMBL Block Renewal' in t:
            _set_paragraph_checkboxes(p, [True])
        elif 'Planning Commission Approval' in t and 'Director' in t:
            _set_paragraph_checkboxes(p, [is_pc, not is_pc])
    _clear_all_red_run_colors(doc)
    # The City file ends with an empty paragraph after the two-column checkbox
    # section, which spills onto a blank second page; drop trailing empties.
    body = doc.element.body
    kids = list(body)
    while len(kids) >= 2 and kids[-2].tag == qn('w:p') and not ''.join(t.text or '' for t in kids[-2].iter(qn('w:t'))).strip() \
            and kids[-2].find(qn('w:pPr') + '/' + qn('w:sectPr')) is None and not list(kids[-2].iter(qn('w:drawing'))):
        body.remove(kids[-2])
        kids = list(body)
    # If the document now ends with an empty final section right after the
    # two-column attachments section, let the attachments section be the final
    # one: Word otherwise reserves a paragraph for the empty section, which
    # overflows onto a blank page.
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
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    prepare_routing_form(doc)
    doc.save(str(out))
    unlock_docx(out)
    return 1


# =============================================================================
# OneMeeting staff report (Director path)
# =============================================================================
def generate_staff_report(workbook_path, template_path, output_path):
    ctx = _get_renewal_context(workbook_path)
    from core.city_templates import render_updated
    if render_updated(template_path, output_path, ctx, workbook_path, 'mbl_renewal', 'staff'): return 1
    doc = Document(template_path)
    area = ctx['Area_Description']
    district = ctx['Council_District'] or 'X'
    subdivision = ctx['Subdivision_Name']
    if subdivision != '(Subdivision Name)' and 'subdivision' not in subdivision.lower():
        subdivision = f"{subdivision} Subdivision"
    for p in doc.paragraphs:
        t = _norm(p.text)
        c = _RunCursor(p)
        if t.startswith('DISTRICT:') and ctx['Council_District']:
            c.set('CHOOSE THE DISTRICT', ctx['Council_District'])
        elif 'Renewal Application No.' in t:
            _replace_runs_contextual(p, exact_map={'XXX': ctx['Application_Renewal_Number']})
        elif t.startswith('(') and 'block of' in t:
            c.set('2800 - 2900 block of Wheeler Street, north side, between Delano Street and Ennis Street', area)
        elif t.startswith('ITEM CAPTION'):
            c.set('ENTER THE CAPTION FROM THE ORDINANCE (see below example)', '')
            c.set('the north side of the 2800-2900 Blocks of Wheeler Street, between Delano Street and Ennis Street ', f"the {ctx['Ordinance_Caption']} ")
            c.set('DISTRICT D - EVANS-SHABAZZ', f"DISTRICT {district} - {ctx['Council_Member']}")
        elif t.startswith('In accordance with'):
            c.set('XX', ctx['Applicant_Property_Address'])
            c.set(' Street', '')
            c.set('(Subdivision Name)', subdivision)
            c.set('XX', area)
            for old in (' block of ', 'XX', ' Street, ', 'X', ' side between ', 'XX', ' Street and ', 'XX'):
                c.set(old, '')
            c.set(' Boulevard as a ', ' as a ')
        elif 'original ordinance' in t:
            c.set('(200X-XXX) ', f"({ctx['Original_Ordinance_No']}) ")
            c.set('20__.', f"{ctx['Original_Ordinance_Year']}.")
        elif t.startswith('As no action is required'):
            c.set('XX ', f"{ctx['Minimum_Building_Line_Text']} ")
            c.span('XXXX', ' Street. ', area + '. ')
        elif 'Your Name, Position, Phone' in t:
            _replace_in_runs(p, {'Your Name, Position, Phone Number': f"{ctx['Planner_Name']}, {ctx['Planner_Position']}, {ctx['Planner_Phone']}"})
    _clear_all_red_run_colors(doc)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    unlock_docx(out)
    return 1


# =============================================================================
# Draft renewal ordinance
# =============================================================================
def generate_draft_ordinance(workbook_path, template_path, map_path, output_path):
    ctx = _get_renewal_context(workbook_path)
    words, num = ctx['Minimum_Building_Line_Words'], ctx['Minimum_Building_Line_Text']
    year = _clean(ctx.get('Ordinance_Year')) or str(datetime.date.today().year)
    with tempfile.TemporaryDirectory() as td:
        staged = Path(td) / 'ordinance_staged.docx'
        _replace_first_content_control_text(template_path, staged, {'MLS Boundary Description': ctx['Ordinance_Caption_Upper'] or 'APPLICATION AREA'})
        doc = Document(staged)
        for p in doc.paragraphs:
            t = p.text.strip()
            if 'City of Houston, Texas, Ordinance No.' in t:
                _set_ordinance_para(p, [(f'City of Houston, Texas, Ordinance No. {year} - ', {'bold': True})])
                pPr = p._p.get_or_add_pPr()
                numPr = pPr.find(qn('w:numPr'))
                if numPr is not None:
                    pPr.remove(numPr)
            elif 'after consideration and evaluation' in t:
                _set_whereas(p, f"the Director, after consideration and evaluation of the application, determined that the prevailing constructed building line of {words} ({num}) feet established by Ordinance No. {ctx['Original_Ordinance_No']} has maintained and preserved the character of the Area and recommends renewal of the ordinance to continue to maintain and preserve the building line character of the SMBLR Block; ", 'NOW, THEREFORE,')
            elif 'by Ordinance No.' in t and 'City Council' in t and 'established' in t:
                _set_whereas(p, f"the City Council of the City of Houston, Texas (\"City\"), by Ordinance No. {ctx['Original_Ordinance_No']} effective on {ctx['Original_Effective_Date']}, established the {ctx['Ordinance_Caption']}, within the City, as a special minimum building line requirement block (\"SMBLR Block\") pursuant to the City's Code of Ordinances (\"Code\"); Ordinance No. {ctx['Original_Ordinance_No']} expired on {ctx['Original_Expiration_Date']}; and")
            elif 'an application was filed' in t:
                _set_whereas(p, f"an application was filed with the City's Department of Planning and Development (\"Department\") on {ctx['Application_Filed_Date']}, accepted on {ctx['Application_Accepted_Date']}, and deemed completed on {ctx['Application_Complete_Date']}, requesting the renewal of the establishment of the minimum building line requirement within the SMBLR Block; and")
            elif 'timely protest' in t:
                if ctx['Approval_Path'] == 'Planning Commission Approval' or ctx['Protest_Count'] > 0:
                    _set_whereas(p, 'a timely protest to the application was filed by a lot owner within the Application Area and the application was referred to the Houston Planning Commission for a public hearing; and')
                else:
                    _set_whereas(p, 'no timely protest to the application has been filed by any lot owner within the Application Area; and')
            elif t.startswith('Section 4.'):
                _set_ordinance_para(p, [
                    ('Section 4. ', {'bold': True}),
                    (f"That the minimum building line of {words} ({num}) feet shall be the building line requirement within the SMBLR Block. The Director is hereby directed to cause a copy of this Ordinance to be filed for recordation in the real property records of Harris County, Texas, where the SMBLR Block is located, as soon as practicable after the effective date of this Ordinance.", {'bold': False}),
                ])
            elif t.startswith('Requested by'):
                _replace_in_runs(p, {'Vonn Tran': ctx['Director_Name']})
            elif t.startswith('L.D. File No.') and _clean(ctx.get('Legal_File_No')):
                _set_ordinance_para(p, [(f"L.D. File No. {_clean(ctx.get('Legal_File_No'))}", {'bold': None})])
        # The City source reserves Exhibit A for Legal.
        if not any('LEAVE BLANK FOR LEGAL' in p.text for p in doc.paragraphs):
            _populate_exhibit_a(doc, ctx)
        _populate_exhibit_b(doc, map_path)
        _populate_exhibit_c(doc, ctx)
        _remove_red_from_doc(doc)
        _normalize_draft_ordinance_fonts(doc)
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        doc.save(out)
        unlock_docx(out)
    _strip_all_content_controls(out)
    _remove_paragraph_section_breaks(out)
    return 1


# =============================================================================
# Step 2 export, RCA packet, HCAD
# =============================================================================
def _detect_step2_renewal_range(excel_path):
    wb = load_workbook(excel_path, data_only=True)
    ws = wb[STEP2_SHEET] if STEP2_SHEET in wb.sheetnames else wb[wb.sheetnames[1]]
    old = _step2_layout(ws) == 'old'
    last_col = 'H' if old else 'F'
    last = 15
    for r in range(15, ws.max_row + 1):
        if _clean(ws.cell(r, 1).value) or _to_int(ws.cell(r, 4).value) > 0:
            last = r
    return f'A1:{last_col}{last}'


def export_step2_staff_analysis_pdf(workbook_path, output_pdf):
    from openpyxl.utils.cell import range_boundaries
    from openpyxl.cell.cell import MergedCell
    excel_path = Path(workbook_path)
    output_pdf = Path(output_pdf)
    tmpdir = Path(tempfile.mkdtemp(prefix='mblren_step2_'))
    tmp_xlsx = tmpdir / 'Renewal_Staff_Analysis.xlsx'
    wb = load_workbook(excel_path, data_only=False)
    wbv = load_workbook(excel_path, data_only=True)
    sheet = STEP2_SHEET if STEP2_SHEET in wb.sheetnames else wb.sheetnames[1]
    ws, wsv = wb[sheet], wbv[sheet]
    area = _detect_step2_renewal_range(excel_path)
    min_col, min_row, max_col, max_row = range_boundaries(area)
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row > max_row or rng.min_col > max_col:
            ws.unmerge_cells(str(rng))
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            if isinstance(cell.value, str) and cell.value.startswith('='):
                cell.value = wsv.cell(r, c).value
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            if r < min_row or r > max_row or c < min_col or c > max_col:
                cell = ws.cell(r, c)
                if not isinstance(cell, MergedCell):
                    cell.value = None
    for sh in list(wb.sheetnames):
        if sh != sheet:
            del wb[sh]
    ws = wb[sheet]
    apply_staff_analysis_page_setup(ws, area, min_row, max_row, min_col, max_col, header_row=14)
    wb.save(tmp_xlsx)
    pdf = _convert_office_to_pdf(tmp_xlsx, tmpdir)
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(pdf, output_pdf)
    return output_pdf


def _split_paths(raw):
    if not raw:
        return []
    if isinstance(raw, str):
        return [x for x in re.split(r'[;|]+', raw) if x.strip()]
    return list(raw)


def generate_rca_packet_pdf(workbook_path, templates, map_path, subdivision_plat_path, tax_map_path,
                            hcad_records_path, original_ordinance_path, output_pdf, include_staff_analysis=False):
    """Routing form, Director or PC approval, boundary map, tax map, [plat,
    staff analysis, HCAD records when supplied], original ordinance."""
    ctx = _get_renewal_context(workbook_path)
    output_pdf = Path(output_pdf)
    from core.packet_work import packet_workspace
    with packet_workspace() as work:
        for label, val in (('boundary map', map_path), ('tax map', tax_map_path), ('original ordinance PDF', original_ordinance_path)):
            if not str(val or '').strip():
                raise ValueError(f'Select the {label} before building the RCA packet.')
        routing = work / '01_Renewal_Routing_Form.docx'
        generate_routing_form(workbook_path, templates['routing'], routing)
        pdfs = [routing_pdf_one_page(_ensure_pdf(routing, work), work)]
        if ctx['Approval_Path'] == 'Planning Commission Approval':
            pc = work / '02_Renewal_PC_Approval.docx'
            generate_pc_approval(workbook_path, templates['pc'], pc)
            pdfs.append(_ensure_pdf(pc, work))
        else:
            da = work / '02_Renewal_Director_Approval.docx'
            generate_director_approval(workbook_path, templates['director'], da)
            pdfs.append(_ensure_pdf(da, work))
        pdfs.append(_ensure_pdf(map_path, work))
        pdfs.append(_ensure_pdf(tax_map_path, work))
        if str(subdivision_plat_path or '').strip():
            pdfs.append(_ensure_pdf(subdivision_plat_path, work))
        if include_staff_analysis:
            staff = work / '05_Renewal_Staff_Analysis.pdf'
            export_step2_staff_analysis_pdf(workbook_path, staff)
            pdfs.append(staff)
        hcad = _split_paths(hcad_records_path)
        if hcad:
            pdfs.extend(condense_hcad_pdfs([_ensure_pdf(p, work) for p in hcad], work))
        pdfs.append(_ensure_pdf(original_ordinance_path, work))
        _merge_pdfs(pdfs, output_pdf)
        return len(pdfs)


def _hcad_collect_records_from_step1(workbook_path):
    """One HCAD property record per normalized account, including separated lots."""
    from core.hcad_identity import collect_hcad_records
    return collect_hcad_records(workbook_path)



def export_hcad_address_list(workbook_path, output_folder):
    wb = load_workbook(workbook_path, data_only=True)
    ci = _get_case_info(wb)
    app_no = re.sub(r'[^A-Za-z0-9_-]+', '', str(ci.get('Application_Number', '') or '')) or 'Case'
    hcad_folder = Path(output_folder) / 'HCAD_Records'
    hcad_folder.mkdir(parents=True, exist_ok=True)
    address_path = hcad_folder / f'MBL_REN_{app_no}_HCAD_Address_List.txt'
    records = _hcad_collect_records_from_step1(workbook_path)
    if not records:
        raise ValueError('No site addresses or TAX_ID values were found on Step 1 for HCAD download.')
    lines = [f"{rec['Address']}\t{rec['TAX_ID']}" if rec['TAX_ID'] else rec['Address'] for rec in records]
    address_path.write_text('\n'.join(lines) + '\n', encoding='utf-8')
    return address_path, len(records)


# =============================================================================
# UI
# =============================================================================
class MBLRenewalWizard(EmbeddedWizardMixin, tk.Frame):
    _boundary_map_wizard_key = 'mbl_renewal'
    def __init__(self, parent=None):
        super().__init__(parent)
        self.title(f'MBL Renewal Planner Wizard {WIZARD_VERSION}')
        self.geometry('1280x820')
        base = sample_input_dir('mbl_renewal')
        self.files = {
            'workbook': str(calculation_template_path('mbl_renewal')),
            'owner_template': str(template_path('mbl_renewal', 'owner_template', 'MBL_REN_ALL_OWNERS.docx')),
            'protest_template': str(template_path('mbl_renewal', 'protest_template', 'MBL_REN_Protest_All_Owners.docx')),
            'reschedule_template': str(template_path('mbl_renewal', 'reschedule_template', 'MBL_REN_Notice_PC_Rescheduling.docx')),
            'acceptance_template': str(template_path('mbl_renewal', 'acceptance_template', 'MBL_REN_Application_Acceptance.docx')),
            'applicant_template': str(template_path('mbl_renewal', 'applicant_template', 'MBL_REN_APPLICANT_COMPLETE_SIGNS.docx')),
            'director_template': str(template_path('mbl_renewal', 'director_template', 'MBL_REN_Directors_Approval.docx')),
            'pc_template': str(template_path('mbl_renewal', 'pc_template', 'MBL_REN_PC_Approval.docx')),
            'routing_template': str(template_path('mbl_renewal', 'routing_template', 'MBL_REN_Routing_Form.docx')),
            'staff_report_template': str(template_path('mbl_renewal', 'staff_report_template', 'MBL_REN_OneMeeting_Staff_Report.docx')),
            'ordinance_template': str(template_path('mbl_renewal', 'ordinance_template', 'MBL_REN_Ordinance_Template.docx')),
            'map': str(base / 'sample_boundary_map.pdf'),
            'tax_map': '',
            'original_ordinance': '',
            'subdivision_plat': '',
            'hcad_records': '',
            'include_staff_analysis': 'no',
            'out': str(Path.home() / 'Documents'),
        }
        for _k, _v in core_session.load('mbl_renewal').items():
            if _k in self.files and not self._is_foreign_bundled_template(_k, _v):
                self.files[_k] = _v
        self._build()

    def _is_foreign_bundled_template(self, key, saved_path):
        if not key.endswith('_template') or not saved_path:
            return False
        try:
            saved = Path(saved_path)
            default = Path(self.files[key])
            if saved.resolve() == default.resolve():
                return False
            return saved.parent.name.lower() == 'mbl_renewal' and saved.parent.parent.name.lower() == 'templates'
        except Exception:
            return False

    def _build(self):
        self.colors = {'nav': '#7C2D12', 'nav2': '#C2410C', 'nav_active': '#EA580C', 'bg': '#F8FAFC', 'card': '#FFFFFF',
                       'border': '#CBD5E1', 'muted': '#64748B', 'text': '#0F172A', 'accent_light': '#FFEDD5', 'accent_mid': '#FDBA74'}
        self.configure(bg=self.colors['bg'])
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        sidebar = tk.Frame(self, bg=self.colors['nav'], width=255)
        sidebar.grid(row=0, column=0, sticky='ns')
        sidebar.grid_propagate(False)
        tk.Label(sidebar, text=f'MBL Renewal\nWizard {WIZARD_VERSION}', bg=self.colors['nav'], fg='white', font=('Segoe UI', 20, 'bold'), justify='left').pack(anchor='nw', padx=22, pady=(24, 8))
        tk.Label(sidebar, text='Document automation for SMBL renewal RCA packets (2026 City templates)', bg=self.colors['nav'], fg='#FFEDD5', font=('Segoe UI', 9), justify='left', wraplength=200).pack(anchor='nw', padx=22, pady=(0, 24))
        self.step_buttons = []
        steps = [('1. Select Files', self._select_files_page), ('2. Review Case Info', self._case_info_page), ('3. Review Owners', self._owners_page), ('4. Generate Documents', self._generate_page), ('5. Peer Reviewer', self._peer_review_page)]
        for i, (label, cmd) in enumerate(steps):
            b = tk.Button(sidebar, text=label, anchor='w', bg=self.colors['nav_active'] if i == 0 else self.colors['nav2'], fg='white', activebackground=self.colors['nav_active'], activeforeground='white', relief='flat', font=('Segoe UI', 10, 'bold' if i == 0 else 'normal'), padx=14, pady=12, command=cmd)
            b.pack(fill='x', padx=12, pady=5)
            self.step_buttons.append(b)
        modules = ('Built modules\n• Case Info (2026 workbook)\n• Owner letters\n• Protest / rescheduling letters\n• Acceptance + applicant letters\n'
                   '• Director / PC approvals\n• 2026 routing form\n• OneMeeting staff report\n• Draft renewal ordinance\n• RCA packet PDF')
        tk.Label(sidebar, text=modules, bg=self.colors['nav'], fg='#FED7AA', font=('Segoe UI', 9), justify='left').pack(side='bottom', anchor='sw', padx=22, pady=22)
        main = tk.Frame(self, bg=self.colors['bg'])
        main.grid(row=0, column=1, sticky='nsew')
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)
        header = tk.Frame(main, bg=self.colors['bg'])
        header.grid(row=0, column=0, sticky='ew', padx=22, pady=(18, 14))
        header.grid_columnconfigure(0, weight=1)
        tk.Label(header, text='MBL Renewal/RCA Document Generator', bg=self.colors['bg'], fg=self.colors['text'], font=('Segoe UI', 25, 'bold')).grid(row=0, column=0, sticky='w')
        tk.Label(header, text='Load one renewal calculation workbook, review Case Info, then generate the City renewal templates and the RCA packet.', bg=self.colors['bg'], fg='#334155', font=('Segoe UI', 12), wraplength=950, justify='left').grid(row=1, column=0, sticky='w', pady=(4, 0))
        tk.Button(header, text='Load Workbook', command=self._case_info_page, bg=self.colors['accent_light'], activebackground=self.colors['accent_mid'], fg=self.colors['text'], font=('Segoe UI', 10, 'bold'), padx=16, pady=10, relief='solid', bd=1).grid(row=0, column=1, rowspan=2, padx=(14, 0), sticky='e')
        _holder = ScrollableContent(main, bg=self.colors['bg'])
        _holder.grid(row=1, column=0, sticky='nsew', padx=22)
        self.content = _holder.interior
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)
        status = tk.LabelFrame(main, text='Status', bg=self.colors['card'], fg=self.colors['text'], font=('Segoe UI', 11, 'bold'), padx=14, pady=10)
        status.grid(row=2, column=0, sticky='ew', padx=22, pady=(8, 14))
        self.status_var = tk.StringVar(value=f'Ready. MBL Renewal Wizard {WIZARD_VERSION} loaded with the 2026 City templates.')
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
        tk.Label(outer, text='Step 1: Select the files for this renewal case', bg=self.colors['bg'], fg=self.colors['text'], font=('Segoe UI', 16, 'bold')).grid(row=0, column=0, columnspan=2, sticky='w', padx=8, pady=(0, 6))
        tk.Label(outer, text='City templates are prefilled from the templates folder. Use Browse only if the City updates a renewal template.', bg=self.colors['bg'], fg=self.colors['muted'], font=('Segoe UI', 10), wraplength=980, justify='left').grid(row=1, column=0, columnspan=2, sticky='w', padx=8, pady=(0, 10))
        setup = self._section_card(outer, 'Required setup', 2, 0, 2)
        self._file_row(setup, 'Renewal calculation workbook', 'workbook', 0)
        self._file_row(setup, 'Boundary map', 'map', 1)
        self._file_row(setup, 'Output folder', 'out', 2, folder=True)
        self._add_boundary_map_controls(setup)
        self._add_application_import_controls(setup)
        templates = self._section_card(outer, 'City templates (2026)', 3, 0)
        rows = [('Owner notification letter', 'owner_template'), ('Protest letter to owners', 'protest_template'),
                ('Notice of PC rescheduling', 'reschedule_template'), ('Application acceptance letter', 'acceptance_template'),
                ('Applicant sign letter', 'applicant_template'), ('Director approval', 'director_template'),
                ('Planning Commission approval', 'pc_template'), ('Routing form', 'routing_template'),
                ('OneMeeting staff report', 'staff_report_template'), ('Renewal draft ordinance', 'ordinance_template')]
        for i, (lab, key) in enumerate(rows):
            self._file_row(templates, lab, key, i)
        att = self._section_card(outer, 'Packet attachments', 3, 1)
        tk.Label(att, text='Upload only what the wizard cannot create: the tax map and the original ordinance. Plat, HCAD records and the Step 2 staff analysis are optional extras.', bg='white', fg=self.colors['muted'], font=('Segoe UI', 9), wraplength=390, justify='left').grid(row=0, column=0, columnspan=3, sticky='w', padx=10, pady=(0, 6))
        self._file_row(att, 'Tax map (required)', 'tax_map', 1)
        self._file_row(att, 'Original ordinance (required)', 'original_ordinance', 2)
        self._file_row(att, 'Subdivision plat (optional)', 'subdivision_plat', 3)
        self._file_row(att, 'HCAD records (optional)', 'hcad_records', 4, multi=True)
        self.staff_var = tk.BooleanVar(value=str(self.files.get('include_staff_analysis', 'no')).lower() in ('yes', 'true', '1'))
        tk.Checkbutton(att, text='Also include the Step 2 staff analysis export in the packet', variable=self.staff_var,
                       command=lambda: self.files.__setitem__('include_staff_analysis', 'yes' if self.staff_var.get() else 'no'),
                       bg='white', fg='#334155', font=('Segoe UI', 9), anchor='w').grid(row=5, column=0, columnspan=3, sticky='w', padx=10, pady=(2, 4))
        hcad_btns = tk.Frame(att, bg='white')
        hcad_btns.grid(row=6, column=0, columnspan=3, sticky='w', padx=10, pady=(2, 4))
        tk.Button(hcad_btns, text='Download HCAD Records', command=self._download_hcad_records, bg=self.colors['accent_light'], activebackground=self.colors['accent_mid'], font=('Segoe UI', 9, 'bold'), padx=12, pady=6, relief='solid', bd=1).pack(side='left', padx=(0, 6))
        tk.Button(hcad_btns, text='Import Downloaded PDFs', command=self._import_hcad_pdfs, bg='#E7E5E4', font=('Segoe UI', 9), padx=12, pady=6).pack(side='left', padx=(0, 6))
        tk.Button(hcad_btns, text='Open HCAD Folder', command=self._open_hcad_folder, bg='#E7E5E4', font=('Segoe UI', 9), padx=12, pady=6).pack(side='left')
        tk.Label(att, text='HCAD: click Download HCAD Records to read Step 1, open Google Chrome for HCAD verification, then save the HCAD PDFs into the HCAD_Records folder inside your output folder and select them for the packet. If HCAD blocks automation, save PDFs there and click Import Downloaded PDFs.', bg='white', fg=self.colors['muted'], font=('Segoe UI', 9), wraplength=390, justify='left').grid(row=7, column=0, columnspan=3, sticky='w', padx=10, pady=(0, 8))

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
            messagebox.showerror('Missing files', 'Select the renewal workbook and output folder first.')
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
        tk.Label(c, text='Renewal Case Info is read from the workbook (Case Info tab plus the live Step 2 values). Required mail-out fields: Letter_Date, Protest_Deadline, Planner_Name, Planner_Phone, Planner_Email.', bg='white', fg='#475569', font=('Segoe UI', 11), wraplength=900, justify='left').grid(row=0, column=0, sticky='w', padx=12, pady=12)
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
            try:
                n = _to_int(ci.get('Protest_Count'))
                path = _clean(ci.get('Approval_Path'))
                if n > 0 and path != 'Planning Commission Approval':
                    tk.Label(snap, text=f'⚠ Protest_Count is {n} but Approval_Path is "{path}". Protested renewals go to the Planning Commission.', bg='white', fg='#B45309', font=('Segoe UI', 9, 'bold'), wraplength=860, justify='left').grid(row=1, column=0, sticky='w', padx=12, pady=(0, 10))
            except Exception:
                pass
        pf = self._section_card(self.content, 'Preflight checks (informational — nothing here blocks generation)', 2, 0)
        build_preflight_panel(pf, run_preflight(self.files.get('workbook'), 'mbl_renewal', self.files.get('map'), self.files.get('out'))).grid(row=0, column=0, sticky='w', padx=12, pady=10)

    def _owners_page(self):
        self._set_active(2)
        self.clear()
        c = self._section_card(self.content, 'Owner review', 0, 0)
        tk.Label(c, text='Owner rows are read from Step 1 GIS data. Preview one letter first, then generate all renewal mail-out letters.', bg='white', fg='#475569', font=('Segoe UI', 11), wraplength=900, justify='left').grid(row=0, column=0, sticky='w', padx=12, pady=12)
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
            core_session.save('mbl_renewal', {k: v for k, v in self.files.items() if isinstance(v, str)})
        except Exception:
            pass

    def _fail(self, e):
        log_path = log_exception('MBL Renewal Planner Wizard', self.files.get('out', ''), e)
        detail = f"\n\nDetails were saved to:\n{log_path}" if log_path else ''
        messagebox.showerror('Error', str(e) + detail)
        self.status_var.set('Error: ' + str(e))

    def _notify_done(self, path):
        self._save_session()
        self.status_var.set(f'Generated: {path}')
        messagebox.showinfo('MBL Renewal Wizard', f'Generated:\n{path}')
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

    def _rca_packet(self):
        try:
            self._check_boundary_map()
            out = self._document_output('MBL_Renewal_RCA_Packet.pdf')
            templates = {'routing': self.files['routing_template'], 'director': self.files['director_template'], 'pc': self.files['pc_template']}
            include_staff = str(self.files.get('include_staff_analysis', 'no')).lower() in ('yes', 'true', '1')
            generate_rca_packet_pdf(self._wb(), templates, self.files.get('map', ''), self.files.get('subdivision_plat', ''), self.files.get('tax_map', ''),
                                    self.files.get('hcad_records', ''), self.files.get('original_ordinance', ''), out, include_staff_analysis=include_staff)
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
        tk.Label(outer, text='Step 4: Generate renewal documents', bg=self.colors['bg'], fg=self.colors['text'], font=('Segoe UI', 16, 'bold')).grid(row=0, column=0, columnspan=2, sticky='w', padx=8, pady=(0, 6))
        btn = dict(bg='#E7E5E4', font=('Segoe UI', 10, 'bold'), padx=18, pady=12, relief='solid', bd=1)
        c = self._section_card(outer, 'Owner Notification Mail-Out', 1, 0, 2)
        tk.Label(c, text='One preview letter or one combined Word document with every owner letter; the boundary map follows each letter.', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=900, justify='left').grid(row=0, column=0, columnspan=2, sticky='w', padx=12, pady=(4, 10))
        tk.Button(c, text='Preview One Renewal Notification Letter', command=lambda: self._letters('owner_template', 'MBL_Renewal_Notification_Preview.docx', 'preview'), **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        tk.Button(c, text='Generate All Renewal Notification Letters', command=lambda: self._letters('owner_template', 'MBL_Renewal_All_Owner_Notification_Letters.docx', 'all'), **btn).grid(row=1, column=1, sticky='w', padx=12, pady=8)
        tk.Label(c, text='Protest filed: the public-hearing letter (Protest_Letter_Date, PC_Hearing_Date). Hearing moved: the notice of PC rescheduling (Reschedule_Letter_Date, new PC_Hearing_Date).', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=900, justify='left').grid(row=2, column=0, columnspan=2, sticky='w', padx=12, pady=(6, 6))
        tk.Button(c, text='Preview Protest Letter', command=lambda: self._letters('protest_template', 'MBL_Renewal_Protest_Letter_Preview.docx', 'preview'), **btn).grid(row=3, column=0, sticky='w', padx=12, pady=8)
        tk.Button(c, text='Generate All Protest Letters', command=lambda: self._letters('protest_template', 'MBL_Renewal_All_Owner_Protest_Letters.docx', 'all'), **btn).grid(row=3, column=1, sticky='w', padx=12, pady=8)
        tk.Button(c, text='Preview Rescheduling Notice', command=lambda: self._letters('reschedule_template', 'MBL_Renewal_Rescheduling_Notice_Preview.docx', 'preview'), **btn).grid(row=4, column=0, sticky='w', padx=12, pady=8)
        tk.Button(c, text='Generate All Rescheduling Notices', command=lambda: self._letters('reschedule_template', 'MBL_Renewal_All_Owner_Rescheduling_Notices.docx', 'all'), **btn).grid(row=4, column=1, sticky='w', padx=12, pady=8)
        app = self._section_card(outer, 'Applicant Letters', 2, 0)
        tk.Label(app, text='Application acceptance letter (Application_Accepted_Date) and the complete / sign-posting letter (Letter_Date).', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=430, justify='left').grid(row=0, column=0, columnspan=2, sticky='w', padx=12, pady=(4, 10))
        tk.Button(app, text='Generate Acceptance Letter', command=lambda: self._single(generate_acceptance_letter, 'acceptance_template', 'MBL_Renewal_Application_Acceptance_Letter.docx'), **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        tk.Button(app, text='Generate Applicant Sign Letter', command=lambda: self._single(generate_applicant_letter, 'applicant_template', 'MBL_Renewal_Applicant_Sign_Letter.docx'), **btn).grid(row=1, column=1, sticky='w', padx=12, pady=8)
        appr = self._section_card(outer, 'Approvals', 2, 1)
        tk.Label(appr, text='Renewal Director Approval (no protest) or Planning Commission Approval (protest filed). For PC approval, enter PC_Noncorner_Lot_Count on the workbook Case Info tab. For an older workbook, add that field in column A and the count in column B. A blank or invalid count prints [ENTER COUNT].', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=430, justify='left').grid(row=0, column=0, columnspan=2, sticky='w', padx=12, pady=(4, 10))
        tk.Button(appr, text='Generate Director Approval', command=lambda: self._single(generate_director_approval, 'director_template', 'MBL_Renewal_Director_Approval.docx'), **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        tk.Button(appr, text='Generate PC Approval', command=lambda: self._single(generate_pc_approval, 'pc_template', 'MBL_Renewal_Planning_Commission_Approval.docx'), **btn).grid(row=1, column=1, sticky='w', padx=12, pady=8)
        forms = self._section_card(outer, 'RCA Packet Forms', 3, 0)
        tk.Label(forms, text='Routing form (checks Building Line and SMBL Block Renewal, fills the previous ordinance, council district, Council deadline) and the OneMeeting staff report.', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=430, justify='left').grid(row=0, column=0, columnspan=2, sticky='w', padx=12, pady=(4, 10))
        tk.Button(forms, text='Generate Routing Form', command=lambda: self._single(generate_routing_form, 'routing_template', 'MBL_Renewal_Routing_Form.docx'), **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        tk.Button(forms, text='Generate OneMeeting Staff Report', command=lambda: self._single(generate_staff_report, 'staff_report_template', 'MBL_Renewal_OneMeeting_Staff_Report.docx'), **btn).grid(row=1, column=1, sticky='w', padx=12, pady=8)
        ordc = self._section_card(outer, 'Draft Ordinance', 3, 1)
        tk.Label(ordc, text='Draft renewal ordinance as a standalone Word document (not inserted into the packet).', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=430, justify='left').grid(row=0, column=0, sticky='w', padx=12, pady=(4, 10))
        tk.Button(ordc, text='Generate Draft Renewal Ordinance', command=lambda: self._single(generate_draft_ordinance, 'ordinance_template', 'MBL_Renewal_Draft_Ordinance.docx', with_map=True), **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        packet = self._section_card(outer, 'RCA Packet PDF Builder', 4, 0, 2)
        tk.Label(packet, text='Routing form, Director or PC approval, boundary map, tax map, original ordinance (plus plat, staff analysis and HCAD records when selected).', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=900, justify='left').grid(row=0, column=0, sticky='w', padx=12, pady=(4, 10))
        tk.Button(packet, text='Generate MBL Renewal RCA Packet PDF', command=self._rca_packet, bg='#FFEDD5', activebackground='#FDBA74', font=('Segoe UI', 10, 'bold'), padx=18, pady=12, relief='solid', bd=1).grid(row=1, column=0, sticky='w', padx=12, pady=8)

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
    print('Start it with Run Wizard.bat, or: python app.py --wizard mbl_renewal')
