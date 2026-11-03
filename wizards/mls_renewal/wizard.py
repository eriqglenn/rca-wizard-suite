"""MLS Renewal Planner Wizard (v3.0, 2026 City templates).

Fills the official City of Houston Special Minimum Lot Size Block *renewal*
templates from the renewal calculation workbook (Step 1 GIS data, Step 2
calculation sheet, Case Info tab) and assembles the RCA packet.

Documents generated (all from the templates in templates/mls_renewal):

    Owner notification letters   SFR / No-SFR mail merge, map after each letter
    Protest letter to owners     public-hearing notice when a protest is filed
    Applicant sign letter        completion / sign-posting letter
    Director approval            renewal criteria form (2026 version)
    Planning Commission approval renewal criteria form
    Routing form                 2026 coversheet routing form (content controls)
    OneMeeting staff report      Director-path or PC-path NOVUS cover letter
    Draft renewal ordinance      standalone Word document
    RCA packet PDF               routing, approval, map, tax map, original ordinance

Every generator opens the City template and replaces only the case-specific
placeholder text inside the existing runs/content controls; no form is rebuilt.
"""
import tkinter as tk
import datetime
import re
from tkinter import filedialog, messagebox
from pathlib import Path
import tempfile, os, shutil, subprocess, sys, zipfile, copy, threading, time, webbrowser, urllib.request
from copy import deepcopy
from collections import Counter
from openpyxl import load_workbook
from core.routing_form import prepare_routing_form, routing_pdf_one_page
from docx import Document
from core.docx_unlock import unlock_docx
from docx.text.paragraph import Paragraph
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import fitz
from lxml import etree

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

WIZARD_VERSION = 'v3.0'
STEP2_SHEET = 'Step 2 - Main Calculation Sheet'
STEP1_SHEET = 'Step 1 - Input Properties Info'
_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'


# =============================================================================
# Value formatting
# =============================================================================
def _format_date(v):
    """City long form, e.g. 'January 18, 2006' (no leading zero, cross-platform)."""
    return f"{v.strftime('%B')} {v.day}, {v.year}"


def _clean(v):
    if v is None:
        return ''
    if isinstance(v, datetime.datetime):
        return _format_date(v)
    if isinstance(v, datetime.date):
        return _format_date(v)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    s = str(v).strip()
    return '' if s.lower() in ('<null>', 'null', 'none') else s


def _to_int(v, default=0):
    try:
        return int(round(float(v)))
    except Exception:
        return default


def _fmt_int(v):
    try:
        return f"{int(round(float(v))):,}"
    except Exception:
        return _clean(v)


def _fmt_pct(v):
    try:
        f = float(v)
        if f <= 1:
            f *= 100
        return f"{round(f):.0f}%"
    except Exception:
        s = _clean(v)
        return s if s.endswith('%') else (s + '%' if s else '')


def _fmt_pct_num(v):
    return _fmt_pct(v).replace('%', '')


_ONES = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten',
         'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen']
_TENS = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety']


def _num_words(n):
    """Counts: 'twenty-one'. Hundreds use the City's 'and' style."""
    try:
        n = int(round(float(n)))
    except Exception:
        return ''
    if n < 0:
        return str(n)
    if n < 20:
        return _ONES[n]
    if n < 100:
        return _TENS[n // 10] + (('-' + _ONES[n % 10]) if n % 10 else '')
    if n < 1000:
        return _ONES[n // 100] + ' hundred' + ((' and ' + _num_words(n % 100)) if n % 100 else '')
    if n < 1000000:
        return _num_words(n // 1000) + ' thousand' + ((' and ' + _num_words(n % 1000)) if n % 1000 else '')
    return str(n)


def _num_words_city(n):
    """Lot sizes in the ordinance style: 'five thousand and seven hundred and seventy-five'."""
    return _num_words(n)


def _cap(s):
    return s[:1].upper() + s[1:] if s else s


def _count_phrase(n):
    """'Sixteen (16)' style count used in the approval findings."""
    return f"{_cap(_num_words(n))} ({n})"


def _parse_date(s):
    """Best-effort parse of a Case Info date string; returns date or None."""
    s = _clean(s)
    if not s:
        return None
    for fmt in ('%B %d, %Y', '%b %d, %Y', '%m/%d/%Y', '%Y-%m-%d', '%B %d %Y'):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except Exception:
            continue
    return None


_EXCEL_ERRORS = {'#NAME?', '#VALUE!', '#REF!', '#N/A', '#DIV/0!', '#NULL!', '#NUM!'}


# =============================================================================
# Workbook readers
# =============================================================================
def _step2_layout(ws):
    """'new' for the 2026 City sheet (D2 = lot size to renew, column F flags),
    'old' for the pre-2026 sheet (E10 most frequent lot size, K15 share)."""
    c2 = _clean(ws['C2'].value).lower()
    if 'lot size to renew' in c2 or _to_int(ws['D2'].value) > 0:
        return 'new'
    d10 = _clean(ws['D10'].value).lower()
    e14 = _clean(ws['E14'].value).lower()
    if 'most frequent' in d10 or 'ranked' in e14 or _to_int(ws['E10'].value) > 0:
        return 'old'
    return 'new'


def _get_step2(wb):
    """Read the 2026 renewal calculation sheet (Step 2 - Main Calculation Sheet).

    Layout: B2 application no., D2 lot size to renew, B3/D3 dates, B4 street,
    D4 block description, B5/D5 cross streets, B6/D6 sides, A9 SFR lots,
    A13 total lots, D9 SFR share, D10 qualification statement, and the property
    table from row 15 (A description, B land use, D original lot size, E current
    lot size, F 'Yes' when the lot is unchanged or meets the lot size)."""
    ws = wb[STEP2_SHEET] if STEP2_SHEET in wb.sheetnames else wb[wb.sheetnames[1]]
    layout = _step2_layout(ws)
    rows = []
    for r in range(15, min(ws.max_row, 140) + 1):
        desc = _clean(ws.cell(r, 1).value)
        if layout == 'old':
            # pre-2026 sheet: D lot size, C corner flag, G 'Y' = within parameters
            cur = _to_int(ws.cell(r, 4).value)
            if not desc and cur <= 0:
                continue
            rows.append({'desc': desc, 'land_use': _clean(ws.cell(r, 2).value).upper(), 'orig': 0,
                         'cur': cur, 'flag': '', 'corner': _clean(ws.cell(r, 3).value).upper() == 'Y'})
            continue
        cur = _to_int(ws.cell(r, 5).value)
        if not desc and cur <= 0:
            continue
        rows.append({
            'desc': desc,
            'land_use': _clean(ws.cell(r, 2).value).upper(),
            'orig': _to_int(ws.cell(r, 4).value),
            'cur': cur,
            'flag': _clean(ws.cell(r, 6).value).upper(),
        })
    if not rows:
        # Stale formula cache (Step 2 mirrors Step 1): read the same raw numbers
        # from Step 1 instead.  No City calculation is re-implemented.
        try:
            s1 = wb[STEP1_SHEET]
            headers = [_clean(s1.cell(1, c).value) for c in range(1, s1.max_column + 1)]
            idx = {h: c + 1 for c, h in enumerate(headers)}
            for r in range(2, min(s1.max_row, 127) + 1):
                num = _clean(s1.cell(r, idx.get('Str_num', 5)).value)
                legal = _clean(s1.cell(r, idx.get('Legal_Dscr_1', 30)).value)
                if not num and not legal:
                    continue
                grp = _clean(s1.cell(r, idx.get('group_dscr', 35)).value).lower()
                lu = 'SFR' if 'single-family' in grp else ('VAC' if 'undeveloped' in grp else (
                    'COM' if 'commercial' in grp else ('MF' if 'multi-family' in grp else 'EXC')))
                rows.append({'desc': f"{num} ({legal})", 'land_use': lu, 'orig': 0,
                             'cur': _to_int(s1.cell(r, idx.get('Total_Land_Area', 41)).value), 'flag': ''})
        except Exception:
            pass

    if layout == 'old':
        # pre-2026 sheet: E10 most frequent lot size, H9 SFR share, F9/A13 totals
        min_lot = _to_int(ws['E10'].value)
        sfr_pct_cell = ws['H9'].value
        total = _to_int(ws['F9'].value) or _to_int(ws['A13'].value) or len(rows)
        qualifies = _clean(ws['E3'].value)
    else:
        min_lot = _to_int(ws['D2'].value)
        sfr_pct_cell = ws['D9'].value
        total = _to_int(ws['A13'].value) or len(rows)
        qualifies = _clean(ws['D10'].value)
    if min_lot <= 0:
        sizes = [x['cur'] for x in rows if x['cur'] > 0]
        min_lot = Counter(sizes).most_common(1)[0][0] if sizes else 0

    sfr_count = _to_int(ws['A9'].value) or sum(1 for x in rows if x['land_use'] == 'SFR')
    try:
        sfr_pct = float(sfr_pct_cell)
    except Exception:
        sfr_pct = (sfr_count / total) if total else 0
    meeting = sum(1 for x in rows if x['flag'] == 'YES')
    if not meeting and rows and not any(x['flag'] for x in rows):
        meeting = sum(1 for x in rows if x['cur'] > 0 and (x['cur'] >= min_lot or (x['orig'] and x['cur'] == x['orig'])))
    meeting_pct = (meeting / total) if total else 0

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
        'Minimum_Lot_Size': min_lot,
        'Minimum_Lot_Size_Text': _fmt_int(min_lot) if min_lot else '',
        'SFR_Count': sfr_count,
        'SFR_Percentage': _fmt_pct(sfr_pct),
        'Total_Lot_Count': total,
        'Lots_Meeting_Minimum': meeting,
        'Lots_Meeting_Percentage': _fmt_pct(meeting_pct),
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
    if 'Case Info' not in wb.sheetnames:
        d = {}
    else:
        ws = wb['Case Info']
        d = {}
        for r in range(1, ws.max_row + 1):
            k = ws.cell(r, 1).value
            if k:
                d[str(k).strip()] = ws.cell(r, 2).value
    # Step 2 is the source of truth for the calculated/auto-read fields: overlay
    # the live Step 2 values so a stale Case Info formula cache cannot win.
    try:
        s2 = _get_step2(wb)
        overlay = {
            'Application_Number': s2.get('Application_Number', ''),
            'Application_Type': 'Special Minimum Lot Size Block Renewal Application',
            'Date_Received': s2.get('Date_Received', ''),
            'Date_Complete': s2.get('Date_Complete', ''),
            'Street_Name': s2.get('Street_Name', ''),
            'Block_Description': s2.get('Block_Description', ''),
            'Side_Of_Street': s2.get('Side_Text', ''),
            'Cross_Streets': (s2.get('Area_Line_2', '') or '').replace('between ', '').strip(),
            'Minimum_Lot_Size': s2.get('Minimum_Lot_Size_Text', ''),
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


def _step1_reader(wb):
    ws = wb[STEP1_SHEET] if STEP1_SHEET in wb.sheetnames else wb[wb.sheetnames[0]]
    headers = [_clean(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]
    idx = {h: c + 1 for c, h in enumerate(headers)}

    def gv(row, name):
        c = idx.get(name)
        return _clean(ws.cell(row, c).value) if c else ''
    return ws, gv


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
            'Property_Lot_Size': gv(r, 'Total_Land_Area') or gv(r, 'Lot Size'),
            'Owner_Mail': owner,
            'Mail_Addr_1': addr1,
            'Mail_Addr_2': gv(r, 'Mail_Addr_2'),
            'Mail_City': gv(r, 'Mail_City'),
            'Mail_State': gv(r, 'Mail_State'),
            'Mail_Zip': gv(r, 'Mail_Zip'),
            'Total_Land_Area': _fmt_int(gv(r, 'Total_Land_Area')),
            'group_dscr': gv(r, 'group_dscr'),
            'Property_Description': gv(r, 'Property Description'),
        })
    from core.lot_divisions import collapse_lot_notifications
    return collapse_lot_notifications(wb,owners)


def _hcad_number(v):
    digits = re.sub(r'\D+', '', v or '')
    return digits.zfill(13) if digits else ''


def _get_property_rows(wb):
    """One row per site address for the ordinance exhibits."""
    ws, gv = _step1_reader(wb)
    rows, seen = [], set()
    for r in range(2, ws.max_row + 1):
        addr = gv(r, 'Site_addr_1')
        if not addr or addr in seen:
            continue
        seen.add(addr)
        unit = gv(r, 'str_unit')
        if unit and unit.upper() not in addr.upper().split():
            addr = f"{addr} {unit}"
        legal_parts = [gv(r, 'Legal_Dscr_1')] + [gv(r, f'Legal_Dscr_{i}') for i in (3, 4)]
        legal = ' '.join(x for x in legal_parts if x)
        rows.append({
            'Address': addr,
            'HCAD': gv(r, 'HCAD_NUM_1') or _hcad_number(gv(r, 'TAX_ID') or gv(r, 'ALT_TAX_ID')),
            'Legal_Description': legal or gv(r, 'Property Description'),
            'Subdivision': gv(r, 'Legal_Dscr_2'),
            'Land_Use': gv(r, 'Land Use') or gv(r, 'group_dscr'),
            'Lot_Size': _fmt_int(gv(r, 'Total_Land_Area') or gv(r, 'Lot Size')),
        })
    return rows


def _get_renewal_context(workbook_path):
    wb = load_workbook(workbook_path, data_only=True)
    ci = _get_case_info(wb)
    s = _get_step2(wb)
    app = s.get('Application_Number') or _clean(ci.get('Application_Number')) or ''
    app = str(app).strip()
    if app.upper().endswith('REN'):
        app_ren, app_prefix = app, app[:-3]
    else:
        app_ren, app_prefix = (f'{app}REN' if app else 'XXREN'), (app or 'XX')
    block = s.get('Block_Description', '')
    side_label = s.get('Side_Label', '')
    c1, c2 = s.get('Cross_Street_1', ''), s.get('Cross_Street_2', '')
    between = f"between {c1} and {c2}" if (c1 or c2) else ''
    area = ', '.join(x for x in [block, side_label, between] if x)
    area_da = f"{block} ({side_label}); {between}" if side_label else f"{block}; {between}"
    caption = (f"{side_label} of the {block}" if side_label else block)
    if between:
        caption += f", {between}"

    min_lot = s.get('Minimum_Lot_Size') or 0
    min_lot_text = s.get('Minimum_Lot_Size_Text') or 'X,XXX'
    total = _to_int(s.get('Total_Lot_Count'))
    meeting = _to_int(s.get('Lots_Meeting_Minimum'))
    protest_n = _to_int(ci.get('Protest_Count'), 0)
    if protest_n <= 0:
        protest_sentence, protest_words, protest_tail = 'No protest was received.', 'No', ' protest was filed.'
    elif protest_n == 1:
        protest_sentence, protest_words, protest_tail = 'One (1) protest was filed.', 'One (1)', ' protest was filed.'
    else:
        protest_words = _count_phrase(protest_n)
        protest_sentence, protest_tail = f'{protest_words} protests were filed.', ' protests were filed.'

    eff = _clean(ci.get('Original_Effective_Date'))
    eff_date = _parse_date(eff)
    hearing = _clean(ci.get('PC_Hearing_Date'))
    hearing_date = _parse_date(hearing)
    planner_phone = _clean(ci.get('Planner_Phone')) or '832-393-XXXX'

    context = {**{k: _clean(v) for k, v in ci.items()}, **s}
    context.update({
        'Application_Number': app or 'XX',
        'Application_Renewal_Number': app_ren,
        'Application_Prefix': app_prefix,
        'Area_Description': area,
        'Area_Description_Upper': area.upper(),
        'Area_Description_DA': area_da,
        'Ordinance_Caption': caption,
        'Ordinance_Caption_Upper': caption.upper(),
        'Between_Streets': between,
        'Applicant_Address_Block': '\n'.join(x for x in [_clean(ci.get('Applicant_Address_1')), _clean(ci.get('Applicant_Address_2'))] if x),
        'Applicant_City_State_Zip': (', '.join(x for x in [_clean(ci.get('Applicant_City')), ' '.join(y for y in [_clean(ci.get('Applicant_State')), _clean(ci.get('Applicant_Zip'))] if y)] if x)).strip(),
        'Total_Lot_Count': str(total) if total else 'XX',
        'Total_Lot_Count_Words': _num_words(total) if total else 'XX',
        'Total_Lot_Phrase': f"{_num_words(total)} ({total})" if total else 'XX (X)',
        'SFR_Count': str(s.get('SFR_Count') or ''),
        'SFR_Percentage_Number': _fmt_pct_num(s.get('SFR_Percentage')) or 'XX',
        'Minimum_Lot_Size_Number': str(min_lot) if min_lot else 'XXXX',
        'Minimum_Lot_Size_Text': min_lot_text,
        'Minimum_Lot_Size_Words': _num_words_city(min_lot) if min_lot else 'X',
        'Qualifying_Lot_Count': str(meeting) if meeting else 'X',
        'Qualifying_Lot_Count_Words': _num_words(meeting) if meeting else 'X',
        'Qualifying_Lot_Count_Words_Cap': _cap(_num_words(meeting)) if meeting else 'X',
        'Qualifying_Lot_Phrase': _count_phrase(meeting) if meeting else 'X (X)',
        'Renewal_Percentage_Number': _fmt_pct_num(s.get('Lots_Meeting_Percentage')) or 'X',
        'Protest_Count': protest_n,
        'Protest_Sentence': protest_sentence,
        'Protest_Words': protest_words,
        'Protest_Tail': protest_tail,
        'Original_Ordinance_No': _clean(ci.get('Original_Ordinance_No')) or 'XXXX-XXX',
        'Original_Effective_Date': eff or 'Month XX, XXXX',
        'Original_Ordinance_Year': str(eff_date.year) if eff_date else (_clean(ci.get('Original_Ordinance_No')).split('-')[0] or 'XXXX'),
        'Original_Expiration_Date': _clean(ci.get('Original_Expiration_Date')) or 'Month XX, XXXX',
        'Application_Filed_Date': _clean(ci.get('Application_Filed_Date')) or _clean(s.get('Date_Received')) or 'Month XX, XXXX',
        'Application_Accepted_Date': _clean(ci.get('Application_Accepted_Date')) or _clean(s.get('Date_Received')) or 'Month XX, XXXX',
        'Application_Complete_Date': _clean(ci.get('Application_Complete_Date')) or _clean(s.get('Date_Complete')) or 'Month XX, XXXX',
        'Council_Deadline': _clean(ci.get('Council_Deadline')) or '',
        'Council_District': _clean(ci.get('Council_District')) or '',
        'Council_Member': _clean(ci.get('Council_Member')).upper() or 'CM NAME',
        'Approval_Path': _clean(ci.get('Approval_Path')) or 'Director Approval',
        'Planner_Name': _clean(ci.get('Planner_Name')) or 'Planner Name',
        'Planner_Position': _clean(ci.get('Planner_Position')) or 'Planner',
        'Planner_Phone': planner_phone,
        'Planner_Email': _clean(ci.get('Planner_Email')) or 'Planner.Name@houstontx.gov',
        'Letter_Date': _clean(ci.get('Letter_Date')) or 'Date',
        'Protest_Letter_Date': _clean(ci.get('Protest_Letter_Date')) or _clean(ci.get('Letter_Date')) or 'Date',
        'Protest_Deadline': _clean(ci.get('Protest_Deadline')) or '[DATE]',
        'PC_Hearing_Date': hearing or '[Date]',
        'PC_Hearing_Weekday': (hearing_date.strftime('%A') + ',') if hearing_date else 'Thursday,',
        'Director_Name': _clean(ci.get('Director_Name')) or 'Vonn Tran',
        'Applicant_Property_Address': _clean(ci.get('Applicant_Property_Address')) or 'XX Street',
        'Subdivision_Name': _clean(ci.get('Subdivision_Name')) or '(Subdivision Name)',
        'Sign_Count_Number': _clean(ci.get('Sign_Count_Number')) or '2',
        'Sign_Count_Words': _clean(ci.get('Sign_Count_Words')) or 'two',
        'Sign_Placement': _clean(ci.get('Sign_Placement')) or 'one sign at each corner',
        'Properties': _get_property_rows(wb),
    })
    return context


# =============================================================================
# Word helpers (run-level text replacement that keeps the template formatting)
# =============================================================================
def _norm(s):
    """Word templates pasted from e-mail carry non-breaking spaces; compare on plain spaces."""
    return (s or '').replace('\xa0', ' ')


def _map_to_png(map_path, temp_dir):
    p = Path(map_path)
    if not p.exists():
        return None
    if p.suffix.lower() in ['.png', '.jpg', '.jpeg']:
        return str(p)
    if p.suffix.lower() == '.pdf':
        doc = fitz.open(str(p))
        page = doc[0]
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        out = Path(temp_dir) / 'map_page.png'
        pix.save(str(out))
        doc.close()
        return str(out)
    return None


def _set_run_text_preserve_xml(run, text):
    """Set visible run text without deleting drawings/shapes in the run.

    Tabs and line breaks in `text` are written as real <w:tab/> / <w:br/>
    elements (the inverse of python-docx's run.text reader), so a template
    tab is never doubled and the City seal drawing anchored inside a run is
    never lost."""
    r = run._r
    visible = r.xpath('.//w:t | .//w:tab | .//w:br | .//w:cr')
    children = list(r)
    insert_at = children.index(visible[0]) if visible and visible[0] in children else len(children)
    for el in visible:
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
    new_els = []
    for li, line in enumerate(str(text).split('\n')):
        if li > 0:
            new_els.append(OxmlElement('w:br'))
        for si, seg in enumerate(line.split('\t')):
            if si > 0:
                new_els.append(OxmlElement('w:tab'))
            if seg != '':
                t = OxmlElement('w:t')
                t.set(qn('xml:space'), 'preserve')
                t.text = seg
                new_els.append(t)
    for offset, el in enumerate(new_els):
        r.insert(insert_at + offset, el)


def _blacken_color_el(color):
    color.set(qn('w:val'), '000000')
    for attr in [qn('w:themeColor'), qn('w:themeTint'), qn('w:themeShade')]:
        if attr in color.attrib:
            del color.attrib[attr]


def _blacken_run(run):
    try:
        run.font.color.rgb = RGBColor(0, 0, 0)
    except Exception:
        pass


def _force_paragraph_color_black(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    rPr = pPr.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        pPr.append(rPr)
    color = rPr.find(qn('w:color'))
    if color is None:
        color = OxmlElement('w:color')
        rPr.append(color)
    _blacken_color_el(color)
    for color_el in paragraph._p.xpath('.//w:rPr/w:color'):
        _blacken_color_el(color_el)


def _has_hyperlink(paragraph):
    return bool(paragraph._p.xpath('.//w:hyperlink'))


def _clear_run_shading(paragraph):
    """Remove run shading / highlight from every run of a paragraph.  The City
    letter templates leave Word's grey field shading (w:shd fill E1E3E6) on
    the «merge field» runs of the address block; the merged address must not
    print with a grey background."""
    for r in paragraph._p.iter(qn('w:r')):
        rpr = r.find(qn('w:rPr'))
        if rpr is None:
            continue
        for tag in ('w:shd', 'w:highlight'):
            for el in rpr.findall(qn(tag)):
                rpr.remove(el)


def _replace_in_runs(paragraph, replacements, allow_collapse=True):
    """Replace placeholder text inside the existing runs of one paragraph.

    Run-level first (formatting untouched).  A placeholder that Word split
    across runs is handled by joining the runs into the first one -- but never
    in a paragraph that contains a hyperlink, because the hyperlink text would
    then end up out of order."""
    changed_any = False
    merge_hit = any(k in _norm(''.join(r.text for r in paragraph.runs)) for k in replacements if '\u00ab' in k)
    for run in paragraph.runs:
        txt = _norm(run.text)
        new_txt = txt
        for k, v in replacements.items():
            if k in new_txt:
                new_txt = new_txt.replace(k, str(v))
        if new_txt != txt:
            changed_any = True
            _set_run_text_preserve_xml(run, new_txt)
            _blacken_run(run)
    if allow_collapse and paragraph.runs and not _has_hyperlink(paragraph):
        full = _norm(''.join(run.text for run in paragraph.runs))
        if any(k in full for k in replacements):
            new = full
            for k, v in replacements.items():
                new = new.replace(k, str(v))
            changed_any = True
            _set_run_text_preserve_xml(paragraph.runs[0], new)
            for run in paragraph.runs[1:]:
                _set_run_text_preserve_xml(run, '')
            _blacken_run(paragraph.runs[0])
    if changed_any:
        _force_paragraph_color_black(paragraph)
        if merge_hit:
            _clear_run_shading(paragraph)
    return changed_any


def _header_footer_parts(doc):
    """The header/footer parts that actually exist in the document.

    python-docx's ``section.header`` / ``section.footer`` silently ADD a new
    header/footer definition (part + reference) when a section has none of that
    type, which would give every merged letter an extra footer.  Walking the
    relationships touches only what the City template already contains."""
    parts = []
    for rel in doc.part.rels.values():
        if rel.is_external:
            continue
        if rel.reltype.endswith('/header') or rel.reltype.endswith('/footer'):
            parts.append(rel.target_part)
    return parts


def _iter_all_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    for part in _header_footer_parts(doc):
        for p_el in part.element.iter(qn('w:p')):
            yield Paragraph(p_el, None)


def _replace_text_nodes(doc, replacements):
    """Low-level pass over every w:t (catches hyperlink text python-docx does
    not expose as runs).  Placeholder colouring is cleared on changed runs."""
    parts = [doc.part] + _header_footer_parts(doc)
    for part in parts:
        for t in part.element.xpath('.//w:t'):
            if not t.text:
                continue
            new = _norm(t.text)
            changed = False
            for k, v in replacements.items():
                if k in new:
                    new = new.replace(k, str(v))
                    changed = True
            if changed:
                t.text = new
                run_el = t.getparent()
                while run_el is not None and run_el.tag != qn('w:r'):
                    run_el = run_el.getparent()
                if run_el is not None:
                    for color_el in run_el.xpath('.//w:color'):
                        _blacken_color_el(color_el)


def _replace_doc(doc, replacements):
    for p in _iter_all_paragraphs(doc):
        _replace_in_runs(p, replacements)
    _replace_text_nodes(doc, replacements)


def _set_exact_para_preserve_first_run(p, text):
    """Set paragraph text into the first run (its formatting kept), blank the rest."""
    if p.runs:
        _set_run_text_preserve_xml(p.runs[0], str(text))
        _blacken_run(p.runs[0])
        for r in p.runs[1:]:
            _set_run_text_preserve_xml(r, '')
    else:
        p.add_run(str(text))
    _force_paragraph_color_black(p)


def _replace_runs_contextual(p, exact_map=None, contains_map=None):
    """Exact-run replacement (normalized text) with a contains fallback."""
    exact_map = exact_map or {}
    contains_map = contains_map or {}
    changed = False
    for r in p.runs:
        txt = _norm(r.text)
        new = txt
        if txt in exact_map:
            new = str(exact_map[txt])
        elif txt.strip() in exact_map:
            lead = txt[:len(txt) - len(txt.lstrip())]
            trail = txt[len(txt.rstrip()):]
            new = lead + str(exact_map[txt.strip()]) + trail
        else:
            for old, val in contains_map.items():
                if old in new:
                    new = new.replace(old, str(val))
        if new != txt:
            _set_run_text_preserve_xml(r, new)
            _blacken_run(r)
            changed = True
    return changed


class _RunCursor:
    """Ordered, position-aware run filling for templates that repeat the same
    placeholder text ('XX', 'X', 'XXXX') several times in one paragraph."""

    def __init__(self, p):
        self.runs = list(p.runs)
        self.i = 0

    def _find(self, text, start):
        for j in range(start, len(self.runs)):
            if _norm(self.runs[j].text) == text:
                return j
        return -1

    def set(self, expected, value):
        j = self._find(expected, self.i)
        if j < 0:
            return False
        _set_run_text_preserve_xml(self.runs[j], str(value))
        _blacken_run(self.runs[j])
        self.i = j + 1
        return True

    def span(self, start, end, value):
        j = self._find(start, self.i)
        if j < 0:
            return False
        k = self._find(end, j + 1)
        if k < 0:
            return False
        _set_run_text_preserve_xml(self.runs[j], str(value))
        _blacken_run(self.runs[j])
        for r in self.runs[j + 1:k + 1]:
            _set_run_text_preserve_xml(r, '')
        self.i = k + 1
        return True


def _clear_all_red_run_colors(doc):
    for p in _iter_all_paragraphs(doc):
        for r in p.runs:
            try:
                if str(r.font.color.rgb).upper() in ('FF0000', 'EE0000', 'C00000'):
                    r.font.color.rgb = RGBColor(0, 0, 0)
            except Exception:
                pass


_PHONE_RE = re.compile(r'\(?832\)?[ -]?393-')


def _fill_phone(p, phone):
    """'at 832-393-' + 'XXXX' (or '(832) 393-' + 'xxxx') -> the planner's phone."""
    runs = p.runs
    for i, r in enumerate(runs):
        txt = _norm(r.text)
        if _PHONE_RE.search(txt) and not re.search(r'393-\d{4}', txt):
            new = _PHONE_RE.sub(phone, txt)
            _set_run_text_preserve_xml(r, new)
            _blacken_run(r)
            for nxt in runs[i + 1:]:
                nt = _norm(nxt.text).strip()
                if nt and set(nt) <= set('Xx'):
                    _set_run_text_preserve_xml(nxt, '')
                elif nt:
                    break
            return True
    return False


_EMAIL_PLACEHOLDERS = ['Planner Name@houstontx.gov', 'Plannner@houstontx.gov', 'Planner.Name@houstontx.gov', 'Name@houstontx.gov']


def _fix_mailto_links(doc, email):
    """Point the template's placeholder mailto: hyperlink at the planner's e-mail."""
    if not email or '@' not in email:
        return
    for rel in list(doc.part.rels.values()):
        try:
            if rel.reltype.endswith('/hyperlink') and rel.is_external:
                tgt = rel.target_ref or ''
                if 'houstontx.gov' in tgt and 'speakercomments' not in tgt:
                    rel._target = f'mailto:{email}'
        except Exception:
            pass


def _map_fit_width(doc, map_png, max_width=Inches(6.5), safety=Inches(0.35)):
    """Width (EMU) at which the map fits inside the text area of the section
    that holds the INSERT MAP paragraph (the last section of the letter).  A
    portrait map taller than the page body would otherwise overflow and push
    a blank page into the letter."""
    try:
        from docx.image.image import Image as _DocxImage
        im = _DocxImage.from_file(map_png)
        aspect = im.px_height / float(im.px_width)
    except Exception:
        return max_width
    try:
        s = doc.sections[-1]
        avail_w = s.page_width - s.left_margin - s.right_margin
        avail_h = s.page_height - s.top_margin - s.bottom_margin - safety
    except Exception:
        return max_width
    width = min(max_width, avail_w)
    if aspect > 0 and width * aspect > avail_h:
        width = avail_h / aspect
    return int(width)


def _trim_after_map(p):
    """Drop empty paragraphs that follow the map paragraph (the City templates
    end the map section with a few blank lines); paragraphs that carry a
    section break, text or a drawing stay."""
    for sib in list(p._p.itersiblings()):
        if sib.tag != qn('w:p'):
            break
        if sib.find(qn('w:pPr') + '/' + qn('w:sectPr')) is not None:
            break
        if ''.join(t.text or '' for t in sib.iter(qn('w:t'))).strip() or list(sib.iter(qn('w:drawing'))):
            break
        sib.getparent().remove(sib)


def _insert_map(doc, map_png):
    if not map_png:
        return
    width = _map_fit_width(doc, map_png)
    for p in doc.paragraphs:
        if 'INSERT MAP' in p.text or 'Insert Map' in p.text:
            p.clear()
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = 1.0
            run = p.add_run()
            run.add_picture(map_png, width=width)
            _trim_after_map(p)
            return
    doc.add_page_break()
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.add_run().add_picture(map_png, width=width)


def _compact_notification_first_page(doc):
    """Only the Spanish notice paragraph is reduced to 10.5 pt so it stays on
    page 1; nothing else in the letter is restyled (brief section 19.3)."""
    spanish_markers = ['¡ATENCIÓN!', 'Si necesita ayuda en español', 'gustaría recibir más información']
    for p in doc.paragraphs:
        text = p.text or ''
        if any(marker in text for marker in spanish_markers):
            for run in p.runs:
                if run.text:
                    run.font.size = Pt(10.5)


# ---------------------------------------------------------------- content controls
def _sdt_is_checkbox(sdt):
    return sdt.find(f'.//{{{_W14}}}checkbox') is not None


def _fill_content_controls(doc, rules):
    """Fill Word plain-text content controls by their placeholder text.

    ``rules`` is a list of (placeholder_text, value, context) where context is
    a substring that must appear in the containing paragraph (or None).  The
    control is replaced by a single run that keeps the placeholder run's
    formatting (minus the red colour) and the content-control wrapper is
    removed, so nothing stale renders on top of the merged text.  Checkbox
    controls are left alone."""
    body = doc.element.body
    for sdt in list(body.iter(qn('w:sdt'))):
        if body not in list(sdt.iterancestors()):
            continue  # already replaced as part of an outer control
        if _sdt_is_checkbox(sdt):
            continue
        content = sdt.find(qn('w:sdtContent'))
        if content is None:
            continue
        txt = _norm(''.join(t.text or '' for t in sdt.iter(qn('w:t')))).strip()
        para = next((a for a in sdt.iterancestors() if a.tag == qn('w:p')), None)
        ptxt = _norm(''.join(t.text or '' for t in para.iter(qn('w:t')))) if para is not None else ''
        value = None
        for key, val, ctx in rules:
            if txt == key and (not ctx or ctx in ptxt):
                value = val
                break
        if value is None:
            continue
        first_run = sdt.find(f'.//{qn("w:r")}')
        new_run = OxmlElement('w:r')
        if first_run is not None and first_run.find(qn('w:rPr')) is not None:
            rpr = deepcopy(first_run.find(qn('w:rPr')))
            for color_el in rpr.findall(qn('w:color')):
                _blacken_color_el(color_el)
            new_run.append(rpr)
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = str(value)
        new_run.append(t)
        parent = sdt.getparent()
        parent.insert(parent.index(sdt), new_run)
        parent.remove(sdt)


def _set_paragraph_checkboxes(p, states):
    """Set the checkbox content controls of one paragraph, in order.
    ``states`` is a list of True/False/None (None = leave as is)."""
    boxes = [s for s in p._p.iter(qn('w:sdt')) if _sdt_is_checkbox(s)]
    for sdt, state in zip(boxes, states):
        if state is None:
            continue
        checked = sdt.find(f'.//{{{_W14}}}checked')
        if checked is not None:
            checked.set(f'{{{_W14}}}val', '1' if state else '0')
        for t in sdt.iter(qn('w:t')):
            t.text = '\u2612' if state else '\u2610'


# ---------------------------------------------------------------- list numbering
def _restart_list_numbering(master, letter):
    """Give the letter's numbered paragraphs (Next Steps) their own list
    instance in the master document so numbering restarts at 1 for every
    owner instead of continuing 3, 4, 5... through the mail-out."""
    try:
        npart = master.part.numbering_part.element
    except Exception:
        return
    nums = npart.findall(qn('w:num'))
    if not nums:
        return
    existing = [int(n.get(qn('w:numId'))) for n in nums if n.get(qn('w:numId'), '').isdigit()]
    next_id = (max(existing) + 1) if existing else 1
    mapping = {}
    for p in letter.paragraphs:
        ppr = p._p.pPr
        if ppr is None:
            continue
        numPr = ppr.find(qn('w:numPr'))
        if numPr is None:
            continue
        nid = numPr.find(qn('w:numId'))
        if nid is None:
            continue
        old = nid.get(qn('w:val'))
        if old not in mapping:
            src = next((n for n in nums if n.get(qn('w:numId')) == old), None)
            if src is None or src.find(qn('w:abstractNumId')) is None:
                continue
            new = OxmlElement('w:num')
            new.set(qn('w:numId'), str(next_id))
            a = OxmlElement('w:abstractNumId')
            a.set(qn('w:val'), src.find(qn('w:abstractNumId')).get(qn('w:val')))
            new.append(a)
            lo = OxmlElement('w:lvlOverride')
            lo.set(qn('w:ilvl'), '0')
            so = OxmlElement('w:startOverride')
            so.set(qn('w:val'), '1')
            lo.append(so)
            new.append(lo)
            tail = npart.find(qn('w:numIdMacAtCleanup'))
            if tail is not None:
                tail.addprevious(new)
            else:
                npart.append(new)
            mapping[old] = str(next_id)
            next_id += 1
        nid.set(qn('w:val'), mapping[old])


# ---------------------------------------------------------------- section combiner
_WT = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _final_sectpr(body):
    last = body[-1] if len(body) else None
    return last if last is not None and last.tag == _WT + 'sectPr' else None


def _append_letter_as_own_section(master, letter):
    """Append one rendered letter to master as its own self-contained section so
    the letterhead footer and margins repeat for every owner (not just the first)."""
    mbody = master.element.body
    lbody = letter.element.body
    msect = _final_sectpr(mbody)
    sect_break_p = master.element.makeelement(_WT + 'p', {})
    ppr = master.element.makeelement(_WT + 'pPr', {})
    sect_break_p.append(ppr)
    # The paragraph that carries the section break renders as one line at the
    # end of the map page; give it no height so a map that fills the page
    # cannot push it (and a blank page) after the letter.
    ppr.append(master.element.makeelement(_WT + 'spacing', {_WT + 'before': '0', _WT + 'after': '0',
                                                            _WT + 'line': '20', _WT + 'lineRule': 'exact'}))
    rpr = master.element.makeelement(_WT + 'rPr', {})
    rpr.append(master.element.makeelement(_WT + 'sz', {_WT + 'val': '2'}))
    rpr.append(master.element.makeelement(_WT + 'szCs', {_WT + 'val': '2'}))
    ppr.append(rpr)
    if msect is not None:
        mbody.remove(msect)
        ppr.append(msect)
    mbody.append(sect_break_p)
    lsect = _final_sectpr(lbody)
    for child in list(lbody):
        if child is lsect:
            continue
        mbody.append(deepcopy(child))
    if lsect is not None:
        mbody.append(deepcopy(lsect))


# =============================================================================
# Owner notification letters (SFR / No-SFR / Protest)
# =============================================================================
def _fill_owner_letter(doc, ctx, owner):
    # Renewal notifications describe the proposed area minimum, not a property area.
    is_protest = 'received a protest' in _norm(' '.join(p.text for p in doc.paragraphs))
    letter_date = ctx['Protest_Letter_Date'] if is_protest else ctx['Letter_Date']
    min_lot = ctx.get('Minimum_Lot_Size_Text') or 'X,XXX'
    area1 = ctx.get('Area_Line_1') or 'XXX block of XX Street, XX and XX sides,'
    area2 = ctx.get('Area_Line_2') or 'between XX and XX Streets'
    addr = owner.get('Mail_Addr_1', '')
    if owner.get('Mail_Addr_2') and owner.get('Mail_Addr_2') not in ('<Null>', '0'):
        addr += '\n' + owner['Mail_Addr_2']
    common = {
        # Also support previously selected renewal templates with the old property token.
        "your property's lot size is": 'the minimum lot size is',
        '«Property_Lot_Size»': min_lot,
        '{{Property_Lot_Size}}': min_lot,
        '«Owner_Mail»': owner.get('Owner_Mail', ''),
        '«Mail_Addr_»': addr,
        '«Mail_Addr_1»': owner.get('Mail_Addr_1', ''),
        '«Mail_Addr_2»': owner.get('Mail_Addr_2', ''),
        '«Mail_City»': owner.get('Mail_City', ''),
        '«Mail_State»': owner.get('Mail_State', ''),
        '«Mail_Zip»': owner.get('Mail_Zip', ''),
        '«Total_Land_Area»': owner.get('Total_Land_Area', ''),
        '«group_dscr»': owner.get('group_dscr', ''),
        'XXX block of XX Street, XX and XX sides,': area1,
        'between XX and XX Streets': area2,
        'between XX and XX Street': area2,
    }
    for p in doc.paragraphs:
        t = _norm(p.text)
        ts = t.strip()
        if ts in ('[DATE]', 'DATE', 'Date'):
            _set_exact_para_preserve_first_run(p, letter_date)
            continue
        if 'no later than' in t and '[DATE]' in t:
            _replace_in_runs(p, {'[DATE]': ctx['Protest_Deadline']})
        if 'public hearing is scheduled' in t:
            _replace_runs_contextual(p, exact_map={'Thursday,': ctx['PC_Hearing_Weekday']},
                                     contains_map={'[Date]': ctx['PC_Hearing_Date']})
        if 'minimum lot size is' in t or 'could not be split into lots smaller than' in t:
            _replace_in_runs(p, {'X, XXX': min_lot, 'X,XXX': min_lot, 'XXXX': min_lot})
        if 'Planner' in t or '393-' in t:
            _replace_runs_contextual(p, exact_map={'Planner’s Name': ctx['Planner_Name'], 'Planner Name': ctx['Planner_Name'],
                                                   "Planner's Name": ctx['Planner_Name']},
                                     contains_map={'Planner’s Name': ctx['Planner_Name'], 'Planner Name': ctx['Planner_Name'],
                                                   'or call the ': 'or call '})
            _fill_phone(p, ctx['Planner_Phone'])
        _replace_in_runs(p, common)
    _replace_text_nodes(doc, {k: ctx['Planner_Email'] for k in _EMAIL_PLACEHOLDERS})
    _fix_mailto_links(doc, ctx['Planner_Email'])


def generate_notification(workbook_path, template_path, map_path, output_path, mode='preview'):
    """Owner mail-out from the SFR, No-SFR or protest letter template.
    mode='preview' renders the first valid owner only; 'all' renders every owner
    as its own section with the map after each letter."""
    ctx = _get_renewal_context(workbook_path)
    wb = load_workbook(workbook_path, data_only=True)
    owners = _get_owners(wb)
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
def generate_applicant_letter(workbook_path, template_path, output_path):
    ctx = _get_renewal_context(workbook_path)
    doc = Document(template_path)
    area1 = (ctx.get('Area_Line_1') or 'XX Block of XX, X side,')
    area2 = (ctx.get('Area_Line_2') or 'between XX and XX').rstrip('.') + '.'
    for p in doc.paragraphs:
        t = _norm(p.text)
        ts = t.strip()
        if ts == 'DATE':
            _set_exact_para_preserve_first_run(p, ctx['Letter_Date'])
        elif ts == 'APPLICANT NAME':
            _set_exact_para_preserve_first_run(p, _clean(ctx.get('Applicant_Name')) or 'Applicant Name')
        elif ts == 'ADDRESS':
            _set_exact_para_preserve_first_run(p, ctx.get('Applicant_Address_Block') or 'Address')
        elif 'Houston, TX 77XXX' in t:
            _replace_in_runs(p, {'Houston, TX 77XXX': ctx.get('Applicant_City_State_Zip') or 'Houston, TX 77XXX'})
        elif 'Block of' in t and 'side' in t:
            _replace_in_runs(p, {'XX Block of XX, X side,': area1})
        elif ts.startswith('between') and 'XX' in t:
            _replace_in_runs(p, {'between XX and XX.': area2})
        elif 'Minimum Lot Size is' in t:
            _replace_runs_contextual(p, exact_map={
                'XX': ctx['Minimum_Lot_Size_Text'],
                'two (2': f"{ctx['Sign_Count_Words']} ({ctx['Sign_Count_Number']}",
                'one': ctx['Sign_Placement'],
            }, contains_map={' sign at each corner.': '.'})
        elif 'feel free to contact' in t:
            _replace_runs_contextual(p, exact_map={'Planner': ctx['Planner_Name'], 'xxxx': '', 'XXXX': ''},
                                     contains_map={' at (832) 393-': f" at {ctx['Planner_Phone']}",
                                                   'Planner Name@houstontx.gov': ctx['Planner_Email']})
    _replace_text_nodes(doc, {k: ctx['Planner_Email'] for k in _EMAIL_PLACEHOLDERS})
    _fix_mailto_links(doc, ctx['Planner_Email'])
    _clear_all_red_run_colors(doc)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    unlock_docx(output_path)
    return 1


# =============================================================================
# Approvals
# =============================================================================
def generate_director_approval(workbook_path, template_path, output_path):
    """2026 Director approval: title number, area line, lot size and original
    ordinance in the finding paragraph, signature name/date."""
    ctx = _get_renewal_context(workbook_path)
    from core.city_templates import render_updated
    if render_updated(template_path, output_path, ctx, workbook_path, 'mls_renewal', 'director'): return 1
    doc = Document(template_path)
    body_paras = set(id(p._p) for p in doc.paragraphs)
    for p in _iter_all_paragraphs(doc):
        t = _norm(p.text)
        ts = t.strip()
        if 'Renewal No.' in t:
            _replace_runs_contextual(p, exact_map={'XX': ctx['Application_Prefix'], 'XXREN': ctx['Application_Renewal_Number']})
        elif id(p._p) in body_paras and ' block of ' in t and 'side' in t and 'between' in t:
            _set_exact_para_preserve_first_run(p, ctx['Area_Description_DA'])
        elif 'minimum lot size of' in t or 'Ordinance No.' in t:
            _replace_runs_contextual(p, exact_map={'XXXX square feet': f"{ctx['Minimum_Lot_Size_Text']} square feet",
                                                   '2005-729': ctx['Original_Ordinance_No']},
                                     contains_map={'XXXX square feet': f"{ctx['Minimum_Lot_Size_Text']} square feet",
                                                   '2005-729': ctx['Original_Ordinance_No']})
        elif 'into effect on' in t:
            _replace_runs_contextual(p, exact_map={'June 14, 2005': ctx['Original_Effective_Date']},
                                     contains_map={'June 14, 2005': ctx['Original_Effective_Date']})
        elif ts == 'Vonn Tran':
            _set_exact_para_preserve_first_run(p, ctx['Director_Name'])
        elif ts == 'Date':
            _set_exact_para_preserve_first_run(p, _clean(ctx.get('Director_Approval_Date')) or 'Date')
    _clear_all_red_run_colors(doc)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    unlock_docx(output_path)
    return 1


def generate_pc_approval(workbook_path, template_path, output_path):
    ctx = _get_renewal_context(workbook_path)
    from core.city_templates import render_updated
    if render_updated(template_path, output_path, ctx, workbook_path, 'mls_renewal', 'pc'): return 1
    doc = Document(template_path)
    for p in _iter_all_paragraphs(doc):
        t = _norm(p.text)
        ts = t.strip()
        if 'Renewal No.' in t:
            _replace_runs_contextual(p, exact_map={'XXX': ctx['Application_Renewal_Number']})
        elif 'The application is for the' in t and 'XX00' in t:
            _replace_runs_contextual(p, exact_map={
                'the': 'the ', 'XX00-XX00': ctx['Area_Description'] + '.', ' block of ': '', 'ABC Street': '', ', ': '',
                'south side, between ABC Avenue, and Live Oak Street.': ''})
        elif 'of the lots are developed' in t:
            _replace_runs_contextual(p, exact_map={'XX%': ctx['SFR_Percentage_Number'] + '%'})
        elif 'protest was filed' in t:
            _replace_runs_contextual(p, exact_map={'X': ctx['Protest_Words']},
                                     contains_map={' protest was filed.': ctx['Protest_Tail']})
        elif 'blockfaces are comprised' in t:
            _replace_runs_contextual(p, exact_map={
                'ten': ctx['Total_Lot_Count_Words'], '10)': ctx['Total_Lot_Count'] + ')',
                'Three (3)': ctx['Qualifying_Lot_Phrase'], '(excluding corner lots) have': 'have',
                'a lot size of': 'a lot size of at least', ' 5,775': ' ' + ctx['Minimum_Lot_Size_Text']})
        elif ts == 'Date':
            _set_exact_para_preserve_first_run(p, _clean(ctx.get('PC_Approval_Date')) or 'Date')
    _clear_all_red_run_colors(doc)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    unlock_docx(output_path)
    return 1


# =============================================================================
# Routing form (2026 coversheet routing form with content controls)
# =============================================================================
def generate_routing_form(workbook_path, template_path, output_path):
    ctx = _get_renewal_context(workbook_path)
    from core.city_templates import render_updated
    if render_updated(template_path, output_path, ctx, workbook_path, 'mls_renewal', 'routing'): return 1
    approval = ctx.get('Approval_Path', 'Director Approval')
    doc = Document(template_path)
    location = f"The {ctx['Area_Description']}" if ctx.get('Block_Description') else \
        'The X block of X Street, X and X sides, between X Street and X Street'
    rules = [
        ('Special Minimum Lot Size Block Renewal Application No. X',
         f"Special Minimum Lot Size Block Renewal Application No. {ctx['Application_Renewal_Number']}", None),
        ('Special Minimum Lot Size/ Block Renewal Application No. X',
         f"Special Minimum Lot Size Block Renewal Application No. {ctx['Application_Renewal_Number']}", None),
        ('The X block of X Street, X and X sides, between X Street and X Street', location, None),
        ('Planner Name, Phone Number', f"{ctx['Planner_Name']}, {ctx['Planner_Phone']}", None),
        ('[DATE]', ctx['Council_Deadline'] or '[DATE]', 'no later than'),
        ('Previous Ordinance: XXXX', f"Previous Ordinance: {ctx['Original_Ordinance_No']}", None),
        # Legal line keeps the short case number ('#137') so the tab gap before
        # 'Lot Size/Building Line Established' survives, as on the City sample.
        ('XXX', ctx['Application_Number'] if ctx['Application_Number'] != 'XX' else 'XXX', 'LOT SIZE'),
        (': XXXX', f": {ctx['Minimum_Lot_Size_Text']}", 'Line Established'),
        ('X', f" {ctx['Council_District']}" if ctx['Council_District'] else ' X', 'Council District'),
    ]
    _fill_content_controls(doc, rules)
    is_pc = approval == 'Planning Commission Approval'
    for i, p in enumerate(doc.paragraphs):
        t = p.text or ''
        if t.startswith('The topic of Request') and i + 1 < len(doc.paragraphs):
            # The parenthetical location wraps; keep every wrapped line aligned
            # with the topic text column (brief section 19.5) instead of the
            # template's first-line-only indent.
            loc = doc.paragraphs[i + 1]
            if loc.text.strip().startswith('('):
                ind = loc._p.get_or_add_pPr().find(qn('w:ind'))
                if ind is not None:
                    left = int(ind.get(qn('w:left'), '0') or 0) + int(ind.get(qn('w:firstLine'), '0') or 0)
                    ind.set(qn('w:left'), str(left))
                    ind.set(qn('w:firstLine'), '0')
                    if qn('w:hanging') in ind.attrib:
                        del ind.attrib[qn('w:hanging')]
        if 'Planning Commission Approval' in t and "Director" in t:
            _set_paragraph_checkboxes(p, [is_pc, not is_pc, True])
        elif 'BUILDING LINE' in t and 'LOT SIZE' in t:
            _set_paragraph_checkboxes(p, [False, True])
        elif 'SMLS Block Renewal' in t and 'SMBL' not in t:
            _set_paragraph_checkboxes(p, [True])
    _clear_all_red_run_colors(doc)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prepare_routing_form(doc)
    doc.save(output_path)
    unlock_docx(output_path)
    return 1


# =============================================================================
# OneMeeting staff report (Director path / Planning Commission path)
# =============================================================================
def generate_staff_report(workbook_path, template_path, output_path):
    ctx = _get_renewal_context(workbook_path)
    from core.city_templates import render_updated
    if render_updated(template_path, output_path, ctx, workbook_path, 'mls_renewal', 'staff'): return 1
    doc = Document(template_path)
    full = _norm(' '.join(p.text for p in doc.paragraphs[:3]))
    is_pc = 'PC APPROVAL' in full.upper()
    area = ctx['Area_Description']
    caption_area = ctx['Ordinance_Caption']
    district = ctx['Council_District'] or 'X'
    phone = ctx['Planner_Phone']
    last4 = phone[8:] if phone.startswith('832-393-') and len(phone) >= 12 else ''
    for p in doc.paragraphs:
        t = _norm(p.text)
        c = _RunCursor(p)
        if not is_pc:
            if t.startswith('DISTRICT:') and ctx['Council_District']:
                c.set('CHOOSE THE DISTRICT', ctx['Council_District'])
            elif 'Renewal Application No.' in t:
                _replace_runs_contextual(p, exact_map={'XXX': ctx['Application_Renewal_Number']})
            elif t.startswith('(') and 'block of' in t:
                c.set('2800 - 2900 block of Wheeler Street, north side, between Delano Street and Ennis Street', area)
            elif t.startswith('ITEM CAPTION'):
                c.set('ENTER THE CAPTION FROM THE ORDINANCE (see below example)', '')
                c.set('the north side of the 2800-2900 Blocks of Wheeler Street, between Delano Street and Ennis Street ', f"the {caption_area} ")
                c.set('DISTRICT D - EVANS-SHABAZZ', f"DISTRICT {district} - {ctx['Council_Member']}")
            elif t.startswith('In accordance with Section 42-197'):
                c.set('XX', ctx['Applicant_Property_Address'])
                c.set(' Street ', ' ')
                c.set('(Subdivision Name) ', f"{ctx['Subdivision_Name']} Subdivision ")
                c.set('XX', area)
                for old in (' block of ', 'XX', ' Street, ', 'X', ' side between ', 'XX', ' Street and ', 'XX'):
                    c.set(old, '')
                c.set(' Boulevard as a ', ' as a ')
            elif 'original ordinance' in t:
                c.set(' (200X-XXX) ', f" ({ctx['Original_Ordinance_No']}) ")
                c.set('200X', ctx['Original_Ordinance_Year'])
            elif t.startswith('As no action is required'):
                c.set('XX ', f"{ctx['Minimum_Lot_Size_Text']} ")
                c.set('XXXX', area)
                for old in (' block of ', 'XXX', ' Street, ', 'X', ' and ', 'X ', 'sides between ', 'XXXX', ' Street and ', 'XXXX'):
                    c.set(old, '')
                c.set(' Street.   ', '.   ')
            elif 'Your Name, Position, Phone Number' in t:
                _replace_in_runs(p, {'Your Name, Position, Phone Number': f"{ctx['Planner_Name']}, {ctx['Planner_Position']}, {phone}"})
        else:
            if 'Renewal Application No.' in t:
                c.set('XXX', ctx['Application_Prefix'])
                c.span('XXXX', ' Street', area)
            elif t.startswith('ORDINANCE renewing'):
                c.span('x', ' Street', caption_area)
                c.set('X', district)
                c.set('CM NAME', ctx['Council_Member'])
            elif t.startswith('In accordance with Section 42-197'):
                c.span('(', ')', ctx['Applicant_Property_Address'])
                c.set('X', ctx['Subdivision_Name'])
                c.span('XXXX', 'Street', area)
                c.set('X', ctx['Protest_Words'])
                many = ctx['Protest_Count'] != 1
                c.set('(s)', 's' if many else '')
                c.set(' was', ' were' if many else ' was')
                c.set('/were', '')
            elif t.startswith('The Houston Planning Commission considered'):
                c.set('Month', ctx['PC_Hearing_Date'] if ctx['PC_Hearing_Date'] != '[Date]' else 'Month X, 202X')
                c.set(' ', '')
                c.set('X', '')
                c.set(', 202', '')
                c.set('X', '')
                c.set('20', ctx['Original_Ordinance_No'])
                c.set('XX', '')
                c.set('-', '')
                c.set('XXX', '')
                c.set('20', ctx['Original_Ordinance_Year'])
                c.set('XX', '')
                c.set('XXXX', ctx['Minimum_Lot_Size_Text'])
                c.span('XXXX', 'Street', area)
            elif 'Council Liaison' in t:
                c.set('Planner Name', ctx['Planner_Name'])
                c.set('Position', ctx['Planner_Position'])
                if last4:
                    c.set('XXXX', last4)
                else:
                    _replace_runs_contextual(p, exact_map={'XXXX': ''}, contains_map={', 832-393-': f", {phone}"})
    _clear_all_red_run_colors(doc)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    unlock_docx(output_path)
    return 1


# =============================================================================
# Draft renewal ordinance
# =============================================================================
def _strip_all_content_controls(docx_path):
    """Unwrap Word content controls (keep their text) in a saved .docx."""
    docx_path = Path(docx_path)
    tmp = Path(tempfile.mkdtemp(prefix='strip_sdt_'))
    with zipfile.ZipFile(docx_path, 'r') as zin:
        zin.extractall(tmp)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    xml_path = tmp / 'word' / 'document.xml'
    if xml_path.exists():
        tree = etree.parse(str(xml_path))
        for sdt in list(tree.xpath('//w:sdt', namespaces=ns)):
            parent = sdt.getparent()
            content = sdt.find('w:sdtContent', namespaces=ns)
            if parent is None or content is None:
                continue
            idx = parent.index(sdt)
            for child in list(content):
                parent.insert(idx, child)
                idx += 1
            parent.remove(sdt)
        tree.write(str(xml_path), xml_declaration=True, encoding='UTF-8', standalone='yes')
    with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for f in tmp.rglob('*'):
            if f.is_file():
                zout.write(f, f.relative_to(tmp))
    shutil.rmtree(tmp)


def _remove_paragraph_section_breaks(docx_path):
    """Remove paragraph-level section breaks that create blank trailing pages."""
    docx_path = Path(docx_path)
    tmp = Path(tempfile.mkdtemp(prefix='no_sect_breaks_'))
    with zipfile.ZipFile(docx_path, 'r') as zin:
        zin.extractall(tmp)
    xml_path = tmp / 'word' / 'document.xml'
    tree = etree.parse(str(xml_path))
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    changed = False
    for sect in list(tree.xpath('//w:pPr/w:sectPr', namespaces=ns)):
        parent = sect.getparent()
        if parent is not None:
            parent.remove(sect)
            changed = True
    if changed:
        tree.write(str(xml_path), xml_declaration=True, encoding='UTF-8', standalone='yes')
        with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for f in tmp.rglob('*'):
                if f.is_file():
                    zout.write(f, f.relative_to(tmp))
    shutil.rmtree(tmp)


def _clone_rpr_from_run(run):
    try:
        rPr = run._r.rPr
        return copy.deepcopy(rPr) if rPr is not None else None
    except Exception:
        return None


def _apply_rpr_to_run(run, rpr):
    if rpr is not None:
        try:
            cur = run._r.rPr
            if cur is not None:
                run._r.remove(cur)
            run._r.insert(0, copy.deepcopy(rpr))
        except Exception:
            pass
    _blacken_run(run)


def _template_run_styles(paragraph):
    """Best bold/body run styles from the existing City template paragraph."""
    bold_rpr = body_rpr = first_rpr = None
    for r in paragraph.runs:
        if not r.text:
            continue
        if first_rpr is None:
            first_rpr = _clone_rpr_from_run(r)
        if bold_rpr is None and (r.bold is True or r.text.strip().upper().startswith(('WHEREAS', 'SECTION'))):
            bold_rpr = _clone_rpr_from_run(r)
        is_red = False
        try:
            is_red = bool(r.font.color.rgb and str(r.font.color.rgb).upper() in ('EE0000', 'FF0000'))
        except Exception:
            pass
        if body_rpr is None and not is_red and r.bold is not True and r.text.strip():
            body_rpr = _clone_rpr_from_run(r)
    return bold_rpr or first_rpr, body_rpr or first_rpr


def _set_ordinance_para(p, segments):
    """Replace the sample text of an existing ordinance paragraph, keeping the
    paragraph formatting and cloning the template's run formatting."""
    bold_rpr, body_rpr = _template_run_styles(p)
    p.clear()
    for text, fmt in segments:
        if not text:
            continue
        r = p.add_run(text)
        _apply_rpr_to_run(r, bold_rpr if fmt.get('bold') is True else body_rpr)
        if fmt.get('bold') is not None:
            r.bold = fmt['bold']
        if fmt.get('italic') is not None:
            r.italic = fmt['italic']
    return p


def _set_whereas(p, body, final_bold=''):
    segments = [('WHEREAS, ', {'bold': True}), (body, {'bold': False})]
    if final_bold:
        segments.append((final_bold, {'bold': True}))
    return _set_ordinance_para(p, segments)


def _remove_red_from_doc(doc):
    for p in _iter_all_paragraphs(doc):
        for r in p.runs:
            _blacken_run(r)


def _delete_paragraph(paragraph):
    el = paragraph._element
    el.getparent().remove(el)
    paragraph._p = paragraph._element = None


def _populate_exhibit_a(doc, ctx):
    props = []
    for x in ctx.get('Properties', []):
        desc = ', '.join(v for v in [x.get('Legal_Description', ''), x.get('Subdivision', '')] if v)
        if desc and desc not in props:
            props.append(desc)
    if not props:
        return
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().upper() == 'PROPERTY DESCRIPTIONS':
            template_p = doc.paragraphs[i + 1] if i + 1 < len(doc.paragraphs) else p
            anchor = p._p
            for desc in props:
                para = copy.deepcopy(template_p._p)
                anchor.addnext(para)
                new_para = Paragraph(para, p._parent)
                bold_rpr, body_rpr = _template_run_styles(new_para)
                new_para.clear()
                r = new_para.add_run(desc)
                _apply_rpr_to_run(r, body_rpr)
                anchor = new_para._p
            break


def _populate_exhibit_b(doc, map_path):
    if not map_path:
        return
    with tempfile.TemporaryDirectory() as td:
        map_png = _map_to_png(map_path, td)
        if not map_png:
            return
        found = False
        to_delete = []
        for p in list(doc.paragraphs):
            txt = p.text.strip()
            if txt == 'BOUNDARY MAP':
                p.clear()
                p.alignment = 1
                p.add_run().add_picture(map_png, width=Inches(6.5))
                found = True
                continue
            if found and txt.startswith('EXHIBIT') and 'C' in txt:
                break
            if found and p._p.xpath('.//w:drawing|.//w:pict'):
                to_delete.append(p)
        for p in to_delete:
            _delete_paragraph(p)


def _clear_cell_keep_format(cell, text):
    if not cell.paragraphs:
        cell.text = text
        return
    p = cell.paragraphs[0]
    bold_rpr, body_rpr = _template_run_styles(p)
    p.clear()
    r = p.add_run(str(text or ''))
    _apply_rpr_to_run(r, body_rpr)
    for extra in list(cell.paragraphs[1:]):
        try:
            extra._element.getparent().remove(extra._element)
        except Exception:
            pass


def _populate_exhibit_c(doc, ctx):
    if not doc.tables:
        return
    table = doc.tables[-1]
    props = ctx.get('Properties', [])
    proto_xml = copy.deepcopy(table.rows[1]._tr if len(table.rows) >= 2 else table.rows[0]._tr)
    while len(table.rows) > 1:
        table._tbl.remove(table.rows[1]._tr)
    for prop in props[:200]:
        tr = copy.deepcopy(proto_xml)
        table._tbl.append(tr)
        cells = table.rows[-1].cells
        vals = [prop.get('Address', ''), prop.get('HCAD', ''), prop.get('Legal_Description', ''), prop.get('Subdivision', '')]
        for cell, val in zip(cells, vals):
            _clear_cell_keep_format(cell, val)


def _normalize_draft_ordinance_fonts(doc):
    """Only fill in missing font metadata on generated runs; never override
    the City template's explicit fonts/sizes."""
    for p in _iter_all_paragraphs(doc):
        for r in p.runs:
            if r.text:
                if r.font.name is None:
                    r.font.name = 'Arial'
                if r.font.size is None:
                    r.font.size = Pt(12)


def _replace_first_content_control_text(docx_path, output_path, replacements):
    """Fill tagged content controls (the ordinance caption) without flattening
    surrounding paragraphs/images."""
    with zipfile.ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}
    xml = files['word/document.xml'].decode('utf-8')
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    root = etree.fromstring(xml.encode('utf-8'))
    for sdt in root.xpath('.//w:sdt', namespaces=ns):
        tag_el = sdt.find('.//w:tag', namespaces=ns)
        alias_el = sdt.find('.//w:alias', namespaces=ns)
        tag = tag_el.get('{%s}val' % ns['w']) if tag_el is not None else ''
        alias = alias_el.get('{%s}val' % ns['w']) if alias_el is not None else ''
        key = tag or alias
        if key in replacements:
            texts = sdt.xpath('.//w:sdtContent//w:t', namespaces=ns)
            if texts:
                texts[0].text = replacements[key]
                for extra in texts[1:]:
                    extra.text = ''
                for r in sdt.xpath('.//w:sdtContent//w:r', namespaces=ns):
                    rPr = r.find('w:rPr', namespaces=ns)
                    if rPr is not None:
                        color = rPr.find('w:color', namespaces=ns)
                        if color is not None:
                            color.set('{%s}val' % ns['w'], '000000')
    files['word/document.xml'] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone='yes')
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)


def generate_draft_ordinance(workbook_path, template_path, map_path, output_path):
    ctx = _get_renewal_context(workbook_path)
    words = ctx['Minimum_Lot_Size_Words']
    num = ctx['Minimum_Lot_Size_Text']
    year = _clean(ctx.get('Ordinance_Year')) or str(datetime.date.today().year)
    with tempfile.TemporaryDirectory() as td:
        staged = Path(td) / 'ordinance_staged.docx'
        _replace_first_content_control_text(template_path, staged, {
            'MLS Boundary Description': ctx.get('Ordinance_Caption_Upper', 'APPLICATION AREA')})
        doc = Document(staged)
        for p in doc.paragraphs:
            t = p.text.strip()
            if 'City of Houston, Texas, Ordinance No.' in t:
                _set_ordinance_para(p, [(f'City of Houston, Texas, Ordinance No. {year} - ', {'bold': True})])
                pPr = p._p.get_or_add_pPr()
                numPr = pPr.find(qn('w:numPr'))
                if numPr is not None:
                    pPr.remove(numPr)
            elif 'by Ordinance No.' in t and 'established' in t:
                _set_whereas(p, f"the City Council of the City of Houston, Texas (\"City\"), by Ordinance No. {ctx['Original_Ordinance_No']} effective on {ctx['Original_Effective_Date']}, established the {ctx['Ordinance_Caption']}, within the City, as a special minimum lot size requirement block (\"SMLSR Block\") pursuant to the City's Code of Ordinances (\"Code\"); Ordinance No. {ctx['Original_Ordinance_No']} expired on {ctx['Original_Expiration_Date']}; and")
            elif 'an application was filed' in t:
                _set_whereas(p, f"an application was filed with the City's Department of Planning and Development (\"Department\") on {ctx['Application_Filed_Date']}, accepted on {ctx['Application_Accepted_Date']}, and deemed completed on {ctx['Application_Complete_Date']}, requesting the renewal of the establishment of the minimum lot size requirement within the SMLSR Block; and")
            elif 'timely protest' in t:
                if ctx.get('Approval_Path') == 'Planning Commission Approval' or ctx.get('Protest_Count', 0) > 0:
                    _set_whereas(p, 'a timely protest to the application was filed by a lot owner within the Application Area and the application was referred to the Houston Planning Commission for a public hearing; and')
                else:
                    _set_whereas(p, 'no timely protest to the application has been filed by any lot owner within the Application Area; and')
            elif 'after consideration and evaluation' in t:
                body = (f"the Director, after consideration and evaluation of the application, determined that the prevailing lot size of at least {words} ({num}) square feet has been maintained since Ordinance No. {ctx['Original_Ordinance_No']} went into effect, and the minimum lot size ordinance has maintained and preserved the character of the Area and recommends renewal of the ordinance to continue to maintain and preserve the lot size character of the SMLSR Block; ")
                _set_whereas(p, body, 'NOW, THEREFORE,')
            elif t.startswith('Section 4.'):
                _set_ordinance_para(p, [
                    ('Section 4. ', {'bold': True}),
                    (f"That the minimum lot size block of {words} ({num}) square feet shall be the lot size requirement within the SMLSR Block. The Director is hereby directed to cause a copy of this Ordinance to be filed for recordation in the real property records of Harris County, Texas, where the SMLSR Block is located, as soon as practicable after the effective date of this Ordinance.", {'bold': False}),
                ])
            elif t.startswith('Requested by'):
                _replace_in_runs(p, {'Vonn Tran': ctx['Director_Name']})
            elif t.startswith('L.D. File No.') and _clean(ctx.get('Legal_File_No')):
                _set_ordinance_para(p, [(f"L.D. File No. {_clean(ctx.get('Legal_File_No'))}", {'bold': None})])
        _populate_exhibit_a(doc, ctx)
        _populate_exhibit_b(doc, map_path)
        _populate_exhibit_c(doc, ctx)
        _remove_red_from_doc(doc)
        _normalize_draft_ordinance_fonts(doc)
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(output_path)
        unlock_docx(output_path)
    _strip_all_content_controls(output_path)
    _remove_paragraph_section_breaks(output_path)
    return 1


# =============================================================================
# PDF helpers, Step 2 export, RCA packet
# =============================================================================
def _find_soffice():
    import shutil as _sh
    for exe in ['soffice', 'libreoffice']:
        found = _sh.which(exe)
        if found:
            return found
    candidates = [
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
        '/usr/bin/soffice', '/usr/bin/libreoffice',
        '/usr/local/bin/soffice', '/usr/local/bin/libreoffice',
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    return None


def _convert_office_to_pdf(input_path, work_dir):
    input_path = Path(input_path)
    work_dir = Path(work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)
    # LibreOffice when installed, otherwise Microsoft Word/Excel (City PCs).
    from core.office_convert import office_to_pdf
    return office_to_pdf(input_path, work_dir)


def _ensure_pdf(input_path, work_dir):
    p = Path(input_path)
    if p.suffix.lower() == '.pdf':
        return p
    if p.suffix.lower() in ['.docx', '.doc', '.xlsx', '.xls']:
        return _convert_office_to_pdf(p, work_dir)
    if p.suffix.lower() in ['.png', '.jpg', '.jpeg']:
        doc = fitz.open()
        img = fitz.open(str(p))
        rect = img[0].rect
        page = doc.new_page(width=rect.width, height=rect.height)
        page.show_pdf_page(rect, img, 0)
        out = Path(work_dir) / (p.stem + '.pdf')
        doc.save(str(out))
        doc.close()
        img.close()
        return out
    raise ValueError(f'Unsupported PDF/attachment type: {p}')


def _detect_step2_renewal_range(excel_path):
    """Meaningful staff-analysis range: header block plus property rows that
    carry a description or a lot size (columns A:F of the 2026 sheet)."""
    wb = load_workbook(excel_path, data_only=True)
    ws = wb[STEP2_SHEET] if STEP2_SHEET in wb.sheetnames else wb[wb.sheetnames[1]]
    old = _step2_layout(ws) == 'old'
    lot_col, last_col = (4, 'H') if old else (5, 'F')
    last = 15
    for r in range(15, ws.max_row + 1):
        desc = _clean(ws.cell(r, 1).value)
        lot = _to_int(ws.cell(r, lot_col).value)
        if desc or lot > 0:
            last = r
    return f'A1:{last_col}{last}'


def export_step2_staff_analysis_pdf(workbook_path, output_pdf):
    from openpyxl.utils.cell import range_boundaries
    excel_path = Path(workbook_path)
    output_pdf = Path(output_pdf)
    tmpdir = Path(tempfile.mkdtemp(prefix='renew_step2_'))
    tmp_xlsx = tmpdir / 'Renewal_Staff_Analysis.xlsx'
    wb = load_workbook(excel_path, data_only=False)
    wbv = load_workbook(excel_path, data_only=True)
    sheet = STEP2_SHEET if STEP2_SHEET in wb.sheetnames else wb.sheetnames[1]
    ws = wb[sheet]
    wsv = wbv[sheet]
    area = _detect_step2_renewal_range(excel_path)
    min_col, min_row, max_col, max_row = range_boundaries(area)
    # The 2026 sheet merges B:C on every property row down to row 97; merged
    # continuation cells are read-only, so unmerge everything outside the
    # export range before clearing it.
    from openpyxl.cell.cell import MergedCell
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


def _pdf_page_is_blank(pdf_path, page_index):
    try:
        doc = fitz.open(str(pdf_path))
        page = doc[page_index]
        pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5), alpha=False)
        data = pix.samples
        nonwhite = 0
        for i in range(0, len(data), pix.n):
            if any(channel < 245 for channel in data[i:i + 3]):
                nonwhite += 1
                if nonwhite > 50:
                    doc.close()
                    return False
        doc.close()
        return True
    except Exception:
        return False


def _merge_pdfs(pdfs, output_pdf):
    try:
        from pypdf import PdfWriter, PdfReader
    except ModuleNotFoundError:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pypdf>=4.0.0'])
        from pypdf import PdfWriter, PdfReader
    writer = PdfWriter()
    for p in pdfs:
        reader = PdfReader(str(p))
        for idx, page in enumerate(reader.pages):
            if _pdf_page_is_blank(p, idx):
                continue
            writer.add_page(page)
    output_pdf = Path(output_pdf)
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    with open(output_pdf, 'wb') as f:
        writer.write(f)
    return output_pdf


def _split_paths(raw):
    if not raw:
        return []
    if isinstance(raw, str):
        return [x for x in re.split(r'[;|]+', raw) if x.strip()]
    return list(raw)


def generate_rca_packet_pdf(workbook_path, templates, map_path, subdivision_plat_path, tax_map_path,
                            hcad_records_path, original_ordinance_path, output_pdf,
                            include_staff_analysis=False):
    """Renewal RCA packet in the 2026 order: routing form, Director or PC
    approval, boundary map, tax map, [plat, staff analysis, HCAD records when
    supplied], original ordinance."""
    ctx = _get_renewal_context(workbook_path)
    output_pdf = Path(output_pdf)
    from core.packet_work import packet_workspace
    with packet_workspace() as work:
        if not str(map_path).strip():
            raise ValueError('Select the boundary map before building the RCA packet.')
        if not str(tax_map_path).strip():
            raise ValueError('Select the tax map before building the RCA packet.')
        if not str(original_ordinance_path).strip():
            raise ValueError('Select the original ordinance PDF before building the RCA packet.')
        routing = work / '01_Renewal_Routing_Form.docx'
        generate_routing_form(workbook_path, templates['routing'], routing)
        pdfs = [routing_pdf_one_page(_ensure_pdf(routing, work), work)]
        if ctx.get('Approval_Path') == 'Planning Commission Approval':
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


# =============================================================================
# HCAD records export (shared pattern with the MLS wizard)
# =============================================================================
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
    address_path = hcad_folder / f'MLS_REN_{app_no}_HCAD_Address_List.txt'
    records = _hcad_collect_records_from_step1(workbook_path)
    if not records:
        raise ValueError('No site addresses or TAX_ID values were found on Step 1 for HCAD download.')
    lines = [f"{rec['Address']}\t{rec['TAX_ID']}" if rec['TAX_ID'] else rec['Address'] for rec in records]
    address_path.write_text('\n'.join(lines) + '\n', encoding='utf-8')
    return address_path, len(records)


# =============================================================================
# UI
# =============================================================================
class MLSRenewalWizard(EmbeddedWizardMixin, tk.Frame):
    _boundary_map_wizard_key = 'mls_renewal'
    def __init__(self, parent=None):
        super().__init__(parent)
        self.title(f'MLS Renewal Planner Wizard {WIZARD_VERSION}')
        self.geometry('1280x820')
        base = sample_input_dir('mls_renewal')
        self.files = {
            'workbook': str(calculation_template_path('mls_renewal')),
            'sfr_template': str(template_path('mls_renewal', 'sfr_template', 'MLS_Renewal_SFR_Letters_to_owners.docx')),
            'no_sfr_template': str(template_path('mls_renewal', 'no_sfr_template', 'MLS_Renewal_NO_SFR_Letters_to_owners.docx')),
            'protest_template': str(template_path('mls_renewal', 'protest_template', 'MLS_REN_Protest_All_Owners.docx')),
            'applicant_template': str(template_path('mls_renewal', 'applicant_template', 'MLS_Renewal_Applicant_Signs.docx')),
            'director_template': str(template_path('mls_renewal', 'director_template', 'MLS_REN_Director_Approval.docx')),
            'pc_template': str(template_path('mls_renewal', 'pc_template', 'MLS_REN_PC_Approval.docx')),
            'routing_template': str(template_path('mls_renewal', 'routing_template', 'MLS_Renewal_Routing_Form.docx')),
            'staff_report_director_template': str(template_path('mls_renewal', 'staff_report_director_template', 'MLS_REN_OneMeeting_Director_Approval.docx')),
            'staff_report_pc_template': str(template_path('mls_renewal', 'staff_report_pc_template', 'MLS_REN_OneMeeting_PC_Approval.docx')),
            'ordinance_template': str(template_path('mls_renewal', 'ordinance_template', 'MLS_RENEWAL_Ordinance_Template.docx')),
            'map': str(base / 'sample_boundary_map.pdf'),
            'tax_map': '',
            'original_ordinance': '',
            'subdivision_plat': '',
            'hcad_records': '',
            'include_staff_analysis': 'no',
            'out': str(Path.home() / 'Documents'),
            'owner_letter_type': 'SFR Renewal Letter',
        }
        for _k, _v in core_session.load('mls_renewal').items():
            if _k in self.files and not self._is_foreign_bundled_template(_k, _v):
                self.files[_k] = _v
        self._build()

    def _is_foreign_bundled_template(self, key, saved_path):
        """True when a remembered template path is another suite version's
        bundled copy (…/templates/mls_renewal/<same file name>).  The session
        file is shared by every installed version, so without this rule a
        newer build would keep merging the previous version's templates.
        A template the planner browsed to elsewhere (e.g. a City download) is
        still restored."""
        if not key.endswith('_template') or not saved_path:
            return False
        try:
            saved = Path(saved_path)
            default = Path(self.files[key])
            if saved.resolve() == default.resolve():
                return False
            return (saved.parent.name.lower() == 'mls_renewal'
                    and saved.parent.parent.name.lower() == 'templates')
        except Exception:
            return False

    def _build(self):
        self.colors = {'nav': '#581C2B', 'nav2': '#7F1D1D', 'nav_active': '#991B1B', 'bg': '#F8FAFC', 'card': '#FFFFFF',
                       'border': '#CBD5E1', 'muted': '#64748B', 'text': '#0F172A', 'accent_light': '#FEE2E2', 'accent_mid': '#FCA5A5'}
        self.configure(bg=self.colors['bg'])
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        sidebar = tk.Frame(self, bg=self.colors['nav'], width=255)
        sidebar.grid(row=0, column=0, sticky='ns')
        sidebar.grid_propagate(False)
        tk.Label(sidebar, text=f'MLS Renewal\nWizard {WIZARD_VERSION}', bg=self.colors['nav'], fg='white', font=('Segoe UI', 20, 'bold'), justify='left').pack(anchor='nw', padx=22, pady=(24, 8))
        tk.Label(sidebar, text='Document automation for SMLS renewal RCA packets (2026 City templates)', bg=self.colors['nav'], fg='#FEE2E2', font=('Segoe UI', 9), justify='left', wraplength=200).pack(anchor='nw', padx=22, pady=(0, 24))
        self.step_buttons = []
        steps = [('1. Select Files', self._select_files_page), ('2. Review Case Info', self._case_info_page), ('3. Review Owners', self._owners_page), ('4. Generate Documents', self._generate_page), ('5. Peer Reviewer', self._peer_review_page)]
        for i, (label, cmd) in enumerate(steps):
            b = tk.Button(sidebar, text=label, anchor='w', bg=self.colors['nav_active'] if i == 0 else self.colors['nav2'], fg='white', activebackground=self.colors['nav_active'], activeforeground='white', relief='flat', font=('Segoe UI', 10, 'bold' if i == 0 else 'normal'), padx=14, pady=12, command=cmd)
            b.pack(fill='x', padx=12, pady=5)
            self.step_buttons.append(b)
        modules = ('Built modules\n• Case Info (2026 workbook)\n• SFR / No-SFR owner letters\n• Protest letter to owners\n• Applicant sign letter\n'
                   '• Director / PC approvals\n• 2026 routing form\n• OneMeeting staff report\n• Draft renewal ordinance\n• RCA packet PDF')
        tk.Label(sidebar, text=modules, bg=self.colors['nav'], fg='#FECACA', font=('Segoe UI', 9), justify='left').pack(side='bottom', anchor='sw', padx=22, pady=22)
        main = tk.Frame(self, bg=self.colors['bg'])
        main.grid(row=0, column=1, sticky='nsew')
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)
        header = tk.Frame(main, bg=self.colors['bg'])
        header.grid(row=0, column=0, sticky='ew', padx=22, pady=(18, 14))
        header.grid_columnconfigure(0, weight=1)
        tk.Label(header, text='MLS Renewal/RCA Document Generator', bg=self.colors['bg'], fg=self.colors['text'], font=('Segoe UI', 25, 'bold')).grid(row=0, column=0, sticky='w')
        tk.Label(header, text='Load one renewal calculation workbook, review Case Info, then generate the City renewal templates and the RCA packet.', bg=self.colors['bg'], fg='#334155', font=('Segoe UI', 12), wraplength=950, justify='left').grid(row=1, column=0, sticky='w', pady=(4, 0))
        tk.Button(header, text='Load Workbook', command=self._case_info_page, bg=self.colors['accent_light'], activebackground=self.colors['accent_mid'], fg=self.colors['text'], font=('Segoe UI', 10, 'bold'), padx=16, pady=10, relief='solid', bd=1).grid(row=0, column=1, rowspan=2, padx=(14, 0), sticky='e')
        _holder = ScrollableContent(main, bg=self.colors['bg'])
        _holder.grid(row=1, column=0, sticky='nsew', padx=22)
        self.content = _holder.interior
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)
        status = tk.LabelFrame(main, text='Status', bg=self.colors['card'], fg=self.colors['text'], font=('Segoe UI', 11, 'bold'), padx=14, pady=10)
        status.grid(row=2, column=0, sticky='ew', padx=22, pady=(8, 14))
        self.status_var = tk.StringVar(value=f'Ready. MLS Renewal Wizard {WIZARD_VERSION} loaded with the 2026 City templates.')
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
        templates = self._section_card(outer, 'City templates', 3, 0)
        tk.Label(templates, text='Select SFR or No-SFR letter type in the dropdown below. Templates are still replaceable.', bg='white', fg=self.colors['muted'], font=('Segoe UI', 9), wraplength=430, justify='left').grid(row=0, column=0, columnspan=3, sticky='w', padx=10, pady=(0, 6))
        tk.Label(templates, text='Owner letter type', bg='white', fg='#334155', font=('Segoe UI', 9, 'bold')).grid(row=1, column=0, sticky='w', padx=10, pady=8)
        self.owner_type_var = tk.StringVar(value=self.files.get('owner_letter_type', 'SFR Renewal Letter'))
        om = tk.OptionMenu(templates, self.owner_type_var, 'SFR Renewal Letter', 'No-SFR Renewal Letter')
        om.grid(row=1, column=1, sticky='ew', padx=8, pady=8)
        rows = [('SFR owner letter', 'sfr_template'), ('No-SFR owner letter', 'no_sfr_template'),
                ('Protest letter to owners', 'protest_template'), ('Applicant sign letter', 'applicant_template'),
                ('Director approval', 'director_template'), ('Planning Commission approval', 'pc_template'),
                ('Routing form', 'routing_template'), ('OneMeeting report (Director path)', 'staff_report_director_template'),
                ('OneMeeting report (PC path)', 'staff_report_pc_template'), ('Renewal draft ordinance', 'ordinance_template')]
        for i, (lab, key) in enumerate(rows, start=2):
            self._file_row(templates, lab, key, i)
        att = self._section_card(outer, 'Packet attachments', 3, 1)
        tk.Label(att, text='The 2026 routing form lists five attachments: routing form, approval, map of the area, tax map and original ordinance. Upload only the two the wizard cannot create; plat and HCAD records are optional extras.', bg='white', fg=self.colors['muted'], font=('Segoe UI', 9), wraplength=390, justify='left').grid(row=0, column=0, columnspan=3, sticky='w', padx=10, pady=(0, 6))
        self._file_row(att, 'Tax map (required)', 'tax_map', 1)
        self._file_row(att, 'Original ordinance (required)', 'original_ordinance', 2)
        self._file_row(att, 'Subdivision plat (optional)', 'subdivision_plat', 3)
        self._file_row(att, 'HCAD records (optional)', 'hcad_records', 4, multi=True)
        self.staff_var = tk.BooleanVar(value=str(self.files.get('include_staff_analysis', 'no')).lower() in ('yes', 'true', '1'))

        def _toggle_staff():
            self.files['include_staff_analysis'] = 'yes' if self.staff_var.get() else 'no'
        tk.Checkbutton(att, text='Also include the Step 2 staff analysis export in the packet', variable=self.staff_var, command=_toggle_staff, bg='white', fg='#334155', font=('Segoe UI', 9), anchor='w').grid(row=5, column=0, columnspan=3, sticky='w', padx=10, pady=(2, 4))
        hcad_btns = tk.Frame(att, bg='white')
        hcad_btns.grid(row=6, column=0, columnspan=3, sticky='w', padx=10, pady=(2, 4))
        tk.Button(hcad_btns, text='Download HCAD Records', command=self._download_hcad_records, bg=self.colors['accent_light'], activebackground=self.colors['accent_mid'], font=('Segoe UI', 9, 'bold'), padx=12, pady=6, relief='solid', bd=1).pack(side='left', padx=(0, 6))
        tk.Button(hcad_btns, text='Import Downloaded PDFs', command=self._import_hcad_pdfs, bg='#E7E5E4', font=('Segoe UI', 9), padx=12, pady=6).pack(side='left', padx=(0, 6))
        tk.Button(hcad_btns, text='Open HCAD Folder', command=self._open_hcad_folder, bg='#E7E5E4', font=('Segoe UI', 9), padx=12, pady=6).pack(side='left')
        tk.Label(att, text='HCAD: click Download HCAD Records to read Step 1, open Google Chrome for HCAD verification, then save the HCAD PDFs into the HCAD_Records folder inside your output folder and select them for the packet. If HCAD blocks automation, save PDFs there and click Import Downloaded PDFs.', bg='white', fg=self.colors['muted'], font=('Segoe UI', 9), wraplength=390, justify='left').grid(row=7, column=0, columnspan=3, sticky='w', padx=10, pady=(0, 8))

    # ---------------- HCAD download workflow (ported from MLS wizard) ----------------
    def _hcad_folder(self):
        base = Path(self.files.get('out') or (Path.home() / 'Documents'))
        return base / 'HCAD_Records'

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
        local = os.environ.get('LOCALAPPDATA')
        pf = os.environ.get('PROGRAMFILES')
        pf86 = os.environ.get('PROGRAMFILES(X86)')
        if local:
            candidates.append(Path(local) / 'Google' / 'Chrome' / 'Application' / 'chrome.exe')
        if pf:
            candidates.append(Path(pf) / 'Google' / 'Chrome' / 'Application' / 'chrome.exe')
        if pf86:
            candidates.append(Path(pf86) / 'Google' / 'Chrome' / 'Application' / 'chrome.exe')
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
        chrome = self._find_chrome_executable()
        try:
            subprocess.Popen([chrome, '--remote-debugging-port=9222', f'--user-data-dir={profile}', '--new-window', 'https://search.hcad.org/'],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            for _ in range(20):
                if self._is_chrome_cdp_available():
                    break
                time.sleep(0.25)
            if show_message:
                messagebox.showinfo('HCAD', 'Google Chrome was opened for HCAD.\n\n'
                                    '1. Complete the HCAD human verification in that Chrome window.\n'
                                    '2. Wait until the HCAD property search page is visible.\n'
                                    '3. Return to this wizard and click OK/Continue.\n\n'
                                    'The wizard will then continue the download automatically.')
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
        messagebox.showinfo('HCAD', f'HCAD will open in Google Chrome for {count} record(s).\n\n'
                            'Complete the human verification in Chrome if prompted.\n'
                            'Do not close the Chrome window.\n\n'
                            'When the HCAD search page is visible, click OK here and the wizard will continue downloading to the HCAD_Records folder.')
        if not self._is_chrome_cdp_available():
            messagebox.showwarning('HCAD', 'The wizard still cannot connect to Chrome.\n\n'
                                   'Make sure the Chrome window opened by the wizard is still open. If Chrome did not open, install Google Chrome '
                                   'or open HCAD manually and place the PDFs in the HCAD_Records folder, then click Import Downloaded PDFs.')
            return

        def worker():
            try:
                result = subprocess.run(core_launch.hcad_command() + ['--file', str(address_file),
                                                                      '--output', str(hcad_folder), '--cdp-url', self._chrome_debug_url(),
                                                                      '--delay-min', '3', '--delay-max', '6'],
                                        cwd=core_launch.hcad_cwd(), creationflags=core_launch.console_flags(), text=True)
                pdfs = sorted(str(p) for p in hcad_folder.glob('*.pdf') if p.is_file())
                report = hcad_folder / 'HCAD_Download_Report.csv'
                selected = ';'.join(pdfs) if pdfs else ''
                if pdfs:
                    msg = f'HCAD download finished. {len(pdfs)} PDF(s) found in {hcad_folder}.'
                else:
                    msg = (f'HCAD download finished, but no PDFs were found in {hcad_folder}.\n\n'
                           'If HCAD blocked the automated download, save the HCAD PDFs manually into that folder and click Import Downloaded PDFs.')
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
            _wb = load_workbook(self._wb(), data_only=True)
            ci = _get_case_info(_wb)
        except Exception as e:
            tk.Label(snap, text=f'Could not read the workbook yet: {e}', bg='white', fg='#B45309', font=('Segoe UI', 9), wraplength=860, justify='left').grid(row=0, column=0, sticky='w', padx=12, pady=10)
        if ci:
            build_snapshot_panel(snap, ci).grid(row=0, column=0, sticky='w', padx=12, pady=10)
            try:
                n = _to_int(ci.get('Protest_Count'))
                path = _clean(ci.get('Approval_Path'))
                if n > 0 and path != 'Planning Commission Approval':
                    tk.Label(snap, text=f'\u26a0 Protest_Count is {n} but Approval_Path is "{path}". Protested renewals go to the Planning Commission.', bg='white', fg='#B45309', font=('Segoe UI', 9, 'bold'), wraplength=860, justify='left').grid(row=1, column=0, sticky='w', padx=12, pady=(0, 10))
            except Exception:
                pass
        pf = self._section_card(self.content, 'Preflight checks (informational \u2014 nothing here blocks generation)', 2, 0)
        build_preflight_panel(pf, run_preflight(self.files.get('workbook'), 'mls_renewal', self.files.get('map'), self.files.get('out'))).grid(row=0, column=0, sticky='w', padx=12, pady=10)

    def _owners_page(self):
        self._set_active(2)
        self.clear()
        c = self._section_card(self.content, 'Owner review', 0, 0)
        tk.Label(c, text='Owner rows are read from Step 1 GIS data. Preview one letter first, then generate all renewal mail-out letters.', bg='white', fg='#475569', font=('Segoe UI', 11), wraplength=900, justify='left').grid(row=0, column=0, sticky='w', padx=12, pady=12)
        try:
            owners = _get_owners(load_workbook(self._wb(), data_only=True))
            tk.Label(c, text=f'{len(owners)} owner letter(s) will be generated.', bg='white', fg=self.colors['text'], font=('Segoe UI', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=12, pady=(0, 6))
            for i, o in enumerate(owners[:60]):
                line = f"{i + 1}. {o['Owner_Mail']} \u2014 {o['Mail_Addr_1']}, {o['Mail_City']} {o['Mail_State']} {o['Mail_Zip']}   ({o['Property_Description']})"
                tk.Label(c, text=line, bg='white', fg='#334155', font=('Segoe UI', 9), anchor='w', wraplength=900, justify='left').grid(row=2 + i, column=0, sticky='w', padx=12, pady=1)
            if len(owners) > 60:
                tk.Label(c, text=f'... and {len(owners) - 60} more.', bg='white', fg=self.colors['muted'], font=('Segoe UI', 9)).grid(row=63, column=0, sticky='w', padx=12, pady=(4, 8))
        except Exception as e:
            tk.Label(c, text=f'Could not read owners yet: {e}', bg='white', fg='#B45309', font=('Segoe UI', 9), wraplength=860, justify='left').grid(row=1, column=0, sticky='w', padx=12, pady=10)

    def _selected_owner_template(self):
        self.files['owner_letter_type'] = self.owner_type_var.get() if hasattr(self, 'owner_type_var') else self.files.get('owner_letter_type', 'SFR Renewal Letter')
        return self.files['sfr_template'] if self.files['owner_letter_type'] == 'SFR Renewal Letter' else self.files['no_sfr_template']

    def _require_map_for_notification(self):
        try:
            self._check_boundary_map()
        except Exception as e:
            messagebox.showerror('Update boundary map', str(e))
            return False
        m = self.files.get('map', '')
        if not m or not Path(m).exists():
            messagebox.showerror('Boundary map required',
                                 'A boundary map is required to generate notification letters, because '
                                 'the map is embedded after each owner letter.\n\n'
                                 'Select a boundary map on the Select Files screen, then try again. '
                                 '(Other documents do not require the map.)')
            self.status_var.set('Notification letters need a boundary map. Select one and retry.')
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
            core_session.save('mls_renewal', {k: v for k, v in self.files.items() if isinstance(v, str)})
        except Exception:
            pass

    def _fail(self, e):
        log_path = log_exception('MLS Renewal Planner Wizard', self.files.get('out', ''), e)
        detail = f"\n\nDetails were saved to:\n{log_path}" if log_path else ''
        messagebox.showerror('Error', str(e) + detail)
        self.status_var.set('Error: ' + str(e))

    def _notify_done(self, path):
        self._save_session()
        self.status_var.set(f'Generated: {path}')
        messagebox.showinfo('MLS Renewal Wizard', f'Generated:\n{path}')
        try:
            os.startfile(str(Path(path).parent))
        except Exception:
            pass

    def _approval_path(self):
        try:
            return _clean(_get_case_info(load_workbook(self._wb(), data_only=True)).get('Approval_Path')) or 'Director Approval'
        except Exception:
            return 'Director Approval'

    # ---------------- generate actions ----------------
    def _preview_notice(self):
        if not self._require_map_for_notification():
            return
        try:
            out = self._document_output('MLS_Renewal_Notification_Preview.docx')
            path = generate_notification(self._wb(), self._selected_owner_template(), self.files.get('map', ''), out, 'preview')
            self._notify_done(path)
        except Exception as e:
            self._fail(e)

    def _all_notice(self):
        if not self._require_map_for_notification():
            return
        try:
            out = self._document_output('MLS_Renewal_All_Owner_Notification_Letters.docx')
            path = generate_notification(self._wb(), self._selected_owner_template(), self.files.get('map', ''), out, 'all')
            self._notify_done(path)
        except Exception as e:
            self._fail(e)

    def _preview_protest(self):
        if not self._require_map_for_notification():
            return
        try:
            out = self._document_output('MLS_Renewal_Protest_Letter_Preview.docx')
            path = generate_notification(self._wb(), self.files['protest_template'], self.files.get('map', ''), out, 'preview')
            self._notify_done(path)
        except Exception as e:
            self._fail(e)

    def _all_protest(self):
        if not self._require_map_for_notification():
            return
        try:
            out = self._document_output('MLS_Renewal_All_Owner_Protest_Letters.docx')
            path = generate_notification(self._wb(), self.files['protest_template'], self.files.get('map', ''), out, 'all')
            self._notify_done(path)
        except Exception as e:
            self._fail(e)

    def _applicant_letter(self):
        try:
            out = self._document_output('MLS_Renewal_Applicant_Sign_Letter.docx')
            generate_applicant_letter(self._wb(), self.files['applicant_template'], out)
            self._notify_done(str(out))
        except Exception as e:
            self._fail(e)

    def _director_approval(self):
        try:
            out = self._document_output('MLS_Renewal_Director_Approval.docx')
            generate_director_approval(self._wb(), self.files['director_template'], out)
            self._notify_done(str(out))
        except Exception as e:
            self._fail(e)

    def _pc_approval(self):
        try:
            out = self._document_output('MLS_Renewal_Planning_Commission_Approval.docx')
            generate_pc_approval(self._wb(), self.files['pc_template'], out)
            self._notify_done(str(out))
        except Exception as e:
            self._fail(e)

    def _routing_form(self):
        try:
            out = self._document_output('MLS_Renewal_Routing_Form.docx')
            generate_routing_form(self._wb(), self.files['routing_template'], out)
            self._notify_done(str(out))
        except Exception as e:
            self._fail(e)

    def _staff_report(self):
        try:
            pc = self._approval_path() == 'Planning Commission Approval'
            tmpl = self.files['staff_report_pc_template'] if pc else self.files['staff_report_director_template']
            out = self._document_output(('MLS_Renewal_OneMeeting_Staff_Report_PC.docx' if pc else 'MLS_Renewal_OneMeeting_Staff_Report.docx'))
            generate_staff_report(self._wb(), tmpl, out)
            self._notify_done(str(out))
        except Exception as e:
            self._fail(e)

    def _draft_ordinance(self):
        try:
            self._check_boundary_map()
            out = self._document_output('MLS_Renewal_Draft_Ordinance.docx')
            generate_draft_ordinance(self._wb(), self.files['ordinance_template'], self.files.get('map', ''), out)
            self._notify_done(str(out))
        except Exception as e:
            self._fail(e)

    def _rca_packet(self):
        try:
            self._check_boundary_map()
            out = self._document_output('MLS_Renewal_RCA_Packet.pdf')
            templates = {'routing': self.files['routing_template'], 'director': self.files['director_template'], 'pc': self.files['pc_template']}
            include_staff = str(self.files.get('include_staff_analysis', 'no')).lower() in ('yes', 'true', '1')
            generate_rca_packet_pdf(self._wb(), templates, self.files.get('map', ''), self.files.get('subdivision_plat', ''),
                                    self.files.get('tax_map', ''), self.files.get('hcad_records', ''),
                                    self.files.get('original_ordinance', ''), out, include_staff_analysis=include_staff)
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
        tk.Label(c, text='Generate one preview letter or one combined Word document with all renewal owner letters (SFR or No-SFR type from Select Files). The selected map is embedded after each letter.', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=900, justify='left').grid(row=0, column=0, columnspan=2, sticky='w', padx=12, pady=(4, 10))
        tk.Button(c, text='Preview One Renewal Notification Letter', command=self._preview_notice, **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        tk.Button(c, text='Generate All Renewal Notification Letters', command=self._all_notice, **btn).grid(row=1, column=1, sticky='w', padx=12, pady=8)
        tk.Label(c, text='Protest filed: the public-hearing letter to every owner uses Protest_Letter_Date and PC_Hearing_Date from Case Info.', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=900, justify='left').grid(row=2, column=0, columnspan=2, sticky='w', padx=12, pady=(6, 6))
        tk.Button(c, text='Preview One Protest Letter', command=self._preview_protest, **btn).grid(row=3, column=0, sticky='w', padx=12, pady=8)
        tk.Button(c, text='Generate All Protest Letters', command=self._all_protest, **btn).grid(row=3, column=1, sticky='w', padx=12, pady=8)
        app = self._section_card(outer, 'Applicant Letter', 2, 0)
        tk.Label(app, text='Generate the renewal applicant complete/sign-posting letter.', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=430, justify='left').grid(row=0, column=0, sticky='w', padx=12, pady=(4, 10))
        tk.Button(app, text='Generate Applicant Sign Letter', command=self._applicant_letter, **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        appr = self._section_card(outer, 'Approvals', 2, 1)
        tk.Label(appr, text='Generate the renewal Director Approval (no protest) or Planning Commission Approval (protest filed).', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=430, justify='left').grid(row=0, column=0, columnspan=2, sticky='w', padx=12, pady=(4, 10))
        tk.Button(appr, text='Generate Director Approval', command=self._director_approval, **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        tk.Button(appr, text='Generate PC Approval', command=self._pc_approval, **btn).grid(row=1, column=1, sticky='w', padx=12, pady=8)
        forms = self._section_card(outer, 'RCA Packet Forms', 3, 0)
        tk.Label(forms, text='Routing form checks Lot Size and SMLS Block Renewal and fills the previous ordinance, council district and Council deadline. The OneMeeting staff report follows Approval_Path (Director or PC template).', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=430, justify='left').grid(row=0, column=0, columnspan=2, sticky='w', padx=12, pady=(4, 10))
        tk.Button(forms, text='Generate Routing Form', command=self._routing_form, **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        tk.Button(forms, text='Generate OneMeeting Staff Report', command=self._staff_report, **btn).grid(row=1, column=1, sticky='w', padx=12, pady=8)
        ordc = self._section_card(outer, 'Draft Ordinance', 3, 1)
        tk.Label(ordc, text='Generate the draft renewal ordinance as a standalone Word document. It is not inserted into the RCA packet.', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=430, justify='left').grid(row=0, column=0, sticky='w', padx=12, pady=(4, 10))
        tk.Button(ordc, text='Generate Draft Renewal Ordinance', command=self._draft_ordinance, **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        packet = self._section_card(outer, 'RCA Packet PDF Builder', 4, 0, 2)
        tk.Label(packet, text='Build the renewal RCA packet PDF in the 2026 order: Routing Form, Director or PC Approval, boundary map, tax map, original ordinance (plus plat, staff analysis and HCAD records when selected).', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=900, justify='left').grid(row=0, column=0, sticky='w', padx=12, pady=(4, 10))
        tk.Button(packet, text='Generate MLS Renewal RCA Packet PDF', command=self._rca_packet, bg='#FEE2E2', activebackground='#FCA5A5', font=('Segoe UI', 10, 'bold'), padx=18, pady=12, relief='solid', bd=1).grid(row=1, column=0, sticky='w', padx=12, pady=8)

        pc = self._section_card(outer, 'Planning Commission Draft Report', 5, 0, 2)
        tk.Label(pc, text='Prepare the renewal PC report and speech from the calculation sheet. The PC map uses the standard MLS colors, symbols, legend and editing controls.', bg='white', fg='#475569', font=('Segoe UI', 10), wraplength=900, justify='left').grid(row=0, column=0, columnspan=3, sticky='w', padx=12, pady=(4, 10))
        tk.Button(pc, text='Generate PC Map', command=lambda: self._open_pc_draft(start_map=True), **btn).grid(row=1, column=0, sticky='w', padx=12, pady=8)
        tk.Button(pc, text='Generate PC Draft Report', command=self._open_pc_draft, **btn).grid(row=1, column=1, sticky='w', padx=12, pady=8)
        tk.Button(pc, text='PC Speech', command=self._open_pc_speech, **btn).grid(row=1, column=2, sticky='w', padx=12, pady=8)

    def _open_pc_draft(self, start_map=False):
        try:
            from wizards.mls.pc_ui import open_pc_workspace
            open_pc_workspace(self, start_map)
        except Exception as error:
            messagebox.showerror('MLS Renewal Planning Commission', str(error), parent=self)

    def _open_pc_speech(self):
        try:
            from wizards.mls.pc_speech_ui import open_pc_speech
            open_pc_speech(self)
        except Exception as error:
            messagebox.showerror('MLS Renewal PC Speech', str(error), parent=self)


if __name__ == '__main__':
    print('This wizard now runs inside the RCA Wizard Suite.')
    print('Start it with Run Wizard.bat, or: python app.py --wizard mls_renewal')
