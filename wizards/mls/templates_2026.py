"""MLS (new application) generators for the 2026 City templates.

The 2026 templates are wording-based (no {{tokens}}): the City's own
placeholder text is replaced inside the existing runs and content controls
through the shared helpers in ``core.docx_merge``.  All case data still comes
from ``mailout_engine`` (Step 1 owners, Step 2 calculations, Case Info), so the
numbers in the letters, approvals and ordinance are unchanged from v14.

Documents:
    Owner notification letter      MLS_Owner_Notification_Letter.docx
    Protest letter to owners       MLS_Protest_All_Owners.docx          (new)
    Notice of public hearing       MLS_Notice_Public_Hearing_Less_51.docx (new, support < 51%)
    Applicant sign letter          MLS_Applicant_Complete_Signs.docx
    Director approval              MLS_Director_Approval.docx
    Routing form                   MLS_Routing Form.docx (content controls)
    OneMeeting staff report        MLS_OneMeeting_Staff_Report.docx      (new)
    Draft ordinance                MLS_Ordinance_Template.docx (content controls)
    RCA packet PDF                 routing, approval, map, tax map (+ optional extras)
"""
from __future__ import annotations

import copy
import datetime
import re
import shutil
import tempfile
from pathlib import Path
from typing import Dict, List

from core.routing_form import prepare_routing_form, routing_pdf_one_page
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from core.docx_merge import (
    _RunCursor, _append_letter_as_own_section, _clean, _clear_all_red_run_colors,
    _compact_notification_first_page, _fill_content_controls, _fill_phone,
    _fix_mailto_links, _insert_map, _iter_all_paragraphs, _map_to_png, _norm,
    _parse_date, _replace_in_runs, _replace_runs_contextual, _replace_text_nodes,
    _restart_list_numbering, _sdt_is_checkbox, _set_exact_para_preserve_first_run,
    _set_paragraph_checkboxes, _set_run_text_preserve_xml, _blacken_run,
    _blacken_color_el, _to_int, _cap,
)
from core.docx_unlock import unlock_docx
from core.pdf_condense import condense_hcad_pdfs

from . import mailout_engine as ME

_EMAIL_PLACEHOLDERS = ['Planner Name@houstontx.gov', 'Plannner@houstontx.gov', 'Planner.Name@houstontx.gov', 'Name@houstontx.gov']


# =============================================================================
# Context
# =============================================================================
def _words(n: int) -> str:
    """Lot sizes in the City's ordinance style: 'six thousand eight hundred and twenty'."""
    ones = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'eleven', 'twelve',
            'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen']
    tens = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety']

    def under100(x):
        if x < 20:
            return ones[x]
        return tens[x // 10] + (('-' + ones[x % 10]) if x % 10 else '')

    def under1000(x, use_and):
        parts = []
        if x >= 100:
            parts.append(ones[x // 100] + ' hundred')
            x %= 100
        if x:
            parts.append(under100(x))
        return (' and ' if use_and and len(parts) == 2 else ' ').join(parts)

    if n < 1000:
        return under1000(n, True)
    if n < 1000000:
        rem = n % 1000
        head = under1000(n // 1000, False) + ' thousand'
        if not rem:
            return head
        if rem < 100:
            return f"{head} and {under100(rem)}"
        return f"{head} {under1000(rem, True)}"
    return str(n)


def _count_words(n: int) -> str:
    return ME.number_word(n) if n <= 50 else ME.number_to_words(n)


def build_context(excel_path) -> Dict[str, str]:
    """Everything the 2026 templates need, on top of the engine's verified
    calculation context (Step 2 numbers, Case Info, ordinance area phrases)."""
    ctx = ME.build_ordinance_context(excel_path)
    ci, _owners = ME.load_case_and_owners(excel_path)
    for k, v in ci.items():
        if v not in (None, '') and k not in ctx:
            ctx[k] = v
    app = _clean(ctx.get('Application_Number')) or ''
    if app in ('0', '0.0'):
        app = ''
    block = _clean(ctx.get('Block_Description'))
    sides = _clean(ctx.get('Side_Of_Street'))
    if sides and re.fullmatch(r'[A-Za-z ]+', sides):
        sides = sides.lower()  # Step 2 may say 'East' / 'West'; the letters read 'east and west sides'
    ctx['Side_Of_Street'] = sides
    cross = _clean(ctx.get('Cross_Streets'))
    if cross and not cross.lower().startswith('between'):
        cross = f"between {cross}"
    area = ', '.join(x for x in [block, sides, cross] if x)
    ordinance_area = ', '.join(x for x in [f"{sides} of the {block}" if (sides and block) else (sides or block), cross] if x)
    ctx['Ordinance_Area_Lower'] = ordinance_area
    min_lot = ME.lot_size_number(ctx.get('Minimum_Lot_Size_Number') or ctx.get('Minimum_Lot_Size') or '')
    try:
        min_int = int(min_lot.replace(',', ''))
    except Exception:
        min_int = 0
    total = _to_int(ctx.get('Total_Lot_Count'))
    qual = _to_int(ctx.get('Qualifying_Lot_Count'), None)
    protest_n = _to_int(ctx.get('Protest_Count'), 0)
    hearing = _clean(ctx.get('PC_Hearing_Date'))
    hearing_date = _parse_date(hearing)
    phone = _clean(ctx.get('Planner_Phone')) or '832-393-XXXX'
    ctx.update({
        'Application_Number': app or 'XXX',
        'Area_Line_1': f"{block}, {sides}," if sides else (f"{block}," if block else 'XX block of XX Street, X and X sides,'),
        'Area_Line_2': cross or 'between XX and XX',
        'Area_Description': area,
        'Area_Description_Period': (area + '.') if area else 'XX block of XX Street, X and X sides, between X Street and X Street.',
        'Ordinance_Area_Lower': ctx.get('Ordinance_Area_Lower') or area,
        'Ordinance_Area_Upper': (ctx.get('Ordinance_Area_Lower') or area).upper(),
        'Minimum_Lot_Size_Text': min_lot or 'XXXX',
        'Minimum_Lot_Size_Words': _words(min_int) if min_int else 'X',
        'Support_Percentage': _clean(ctx.get('Support_Percentage')) or 'XX',
        'SFR_Percentage': _clean(ctx.get('SFR_Percentage')) or 'XX',
        'Total_Lot_Count': str(total) if total else 'XX',
        'Total_Lot_Count_Words': _count_words(total) if total else 'XX',
        'Qualifying_Lot_Count': str(qual) if qual is not None else 'X',
        'Qualifying_Lot_Count_Words': _count_words(qual) if qual is not None else 'X',
        'Protest_Count': protest_n,
        'Letter_Date': _clean(ctx.get('Letter_Date')) or 'Date',
        'Protest_Letter_Date': _clean(ctx.get('Protest_Letter_Date')) or _clean(ctx.get('Letter_Date')) or 'Date',
        'Protest_Deadline': _clean(ctx.get('Protest_Deadline')) or '[DATE]',
        'Land_Use_Deadline': _clean(ctx.get('Land_Use_Deadline')) or _clean(ctx.get('Protest_Deadline')) or '[DATE]',
        'PC_Hearing_Date': hearing or '[Date]',
        'PC_Hearing_Weekday': (hearing_date.strftime('%A') + ',') if hearing_date else 'Thursday,',
        'Planner_Name': _clean(ctx.get('Planner_Name')) or 'Planner Name',
        'Planner_Position': _clean(ctx.get('Planner_Position')) or 'Planner',
        'Planner_Phone': phone,
        'Planner_Email': _clean(ctx.get('Planner_Email')) or 'Planner.Name@houstontx.gov',
        'Council_District': _clean(ctx.get('Council_District')) or '',
        'Council_Member': _clean(ctx.get('Council_Member')).upper() or 'CM NAME',
        'Council_Deadline': _clean(ctx.get('Council_Deadline')) or '',
        'Director_Name': _clean(ctx.get('Director_Name')) or 'Vonn Tran',
        'Applicant_Name': _clean(ctx.get('Applicant_Name')) or 'Applicant Name',
        'Applicant_Address_Block': '\n'.join(x for x in [_clean(ctx.get('Applicant_Address_1')), _clean(ctx.get('Applicant_Address_2'))] if x) or 'Address',
        'Applicant_City_State_Zip': (', '.join(x for x in [_clean(ctx.get('Applicant_City')), ' '.join(y for y in [_clean(ctx.get('Applicant_State')), _clean(ctx.get('Applicant_Zip'))] if y)] if x)).strip() or 'Houston, TX XXXXX',
        'Applicant_Property_Address': _clean(ctx.get('Applicant_Property_Address')) or 'XXXX XXX Street',
        'Subdivision_Name': _clean(ctx.get('Subdivision_Name')) or 'SUBDIVISION NAME',
        'Sign_Count_Number': _clean(ctx.get('Sign_Count_Number')) or '2',
        'Sign_Count_Words': _clean(ctx.get('Sign_Count_Words')) or 'two',
        'Sign_Placement': _clean(ctx.get('Sign_Placement')) or 'one sign at each corner',
        'Approval_Path': _clean(ctx.get('Approval_Path')) or 'Director Approval',
    })
    ctx['Qualifying_Phrase'] = f"{ctx['Qualifying_Lot_Count_Words']} ({ctx['Qualifying_Lot_Count']}) out of {ctx['Total_Lot_Count_Words']} ({ctx['Total_Lot_Count']})"
    ctx['Lot_Size_Phrase'] = f"{ctx['Minimum_Lot_Size_Words']} ({ctx['Minimum_Lot_Size_Text']})"
    return ctx


def _owner_legal(owner: Dict[str, str]) -> str:
    l1 = _clean(owner.get('Legal_Dscr_1'))
    l2 = _clean(owner.get('Legal_Dscr_2'))
    if l1.rstrip().endswith('&') and l2:
        return f"{l1} {l2}"
    return l1


# =============================================================================
# Owner letters (notification / protest / notice of public hearing)
# =============================================================================
def _fill_owner_letter(doc, ctx, owner):
    # The notification describes the proposed block minimum, not a property area.
    full = _norm(' '.join(p.text for p in doc.paragraphs))
    is_protest = 'received a protest' in full
    letter_date = ctx['Protest_Letter_Date'] if is_protest else ctx['Letter_Date']
    min_lot = ctx['Minimum_Lot_Size_Text']
    addr2 = _clean(owner.get('Mail_Addr_2'))
    addr2 = '' if addr2 in ('<Null>', '0') else addr2
    common = {
        "Preliminary staff review of the application indicates that your property's lot size is":
            'A preliminary review indicates the proposed minimum lot size for this block would be',
        '«Property_Lot_Size»': min_lot,
        '{{Property_Lot_Size}}': min_lot,
        '«OWNER_MAIL_To»': owner.get('Mailto', ''),
        '«Mailto»': owner.get('Mailto', ''),
        '«Mail_Addr_1» «Mail_Addr_2»': owner.get('Mail_Addr_1', '') + (('\n' + addr2) if addr2 else ''),
        '«Mail_Addr_1»«Mail_Addr_2»': owner.get('Mail_Addr_1', '') + (('\n' + addr2) if addr2 else ''),
        '«Mail_Addr_1»': owner.get('Mail_Addr_1', ''),
        '«Mail_Addr_2»': addr2,
        '«Mail_City»': owner.get('Mail_City', ''),
        '«Mail_State»': owner.get('Mail_State', ''),
        '«Mail_Zip»': owner.get('Mail_Zip', ''),
        '«Legal_Dscr_1»': _owner_legal(owner),
        '«group_dscr»': owner.get('group_dscr', ''),
        'x block of x Street, north and south sides,': ctx['Area_Line_1'],
        'between x Avenue and x Street': ctx['Area_Line_2'],
        'X and X block of XX, X side,': ctx['Area_Line_1'],
        'between XX and XX': ctx['Area_Line_2'],
    }
    for p in doc.paragraphs:
        t = _norm(p.text)
        ts = t.strip()
        if ts in ('[Date]', '[DATE]', 'Date', 'DATE'):
            _set_exact_para_preserve_first_run(p, letter_date)
            continue
        if ts.startswith('X and X block of XX') and 'between' in ts:
            # notification letter: both subject lines live in one paragraph;
            # make the wrapped line sit under the first (template uses a
            # first-line indent only).
            _set_exact_para_preserve_first_run(p, f"{ctx['Area_Line_1']}\n{ctx['Area_Line_2']}")
            ind = p._p.get_or_add_pPr().find(qn('w:ind'))
            if ind is not None:
                left = int(ind.get(qn('w:left'), '0') or 0) + int(ind.get(qn('w:firstLine'), '0') or 0)
                ind.set(qn('w:left'), str(left))
                ind.set(qn('w:firstLine'), '0')
                if qn('w:hanging') in ind.attrib:
                    del ind.attrib[qn('w:hanging')]
            continue
        if 'land use designation' in t and '[DATE]' in t:
            _replace_in_runs(p, {'[DATE]': ctx['Land_Use_Deadline']})
        elif 'no later than' in t and '[DATE]' in t:
            _replace_in_runs(p, {'[DATE]': ctx['Protest_Deadline']})
        if 'public hearing is scheduled' in t or 'public hearing will be held' in t:
            _replace_runs_contextual(p, exact_map={'Thursday,': ctx['PC_Hearing_Weekday']},
                                     contains_map={'[Date]': ctx['PC_Hearing_Date']})
            _replace_in_runs(p, {'[Date]': ctx['PC_Hearing_Date'], 'X%': f"{ctx['Support_Percentage']}%",
                                 'x square feet': f"{min_lot} square feet"})
        if 'square feet' in t:
            # Word splits the bold red number into several runs in the City
            # template. Replace only its text span, retaining the paragraph's styling.
            from core.property_letters import replace_span
            text = ''.join(node.text or '' for node in p._p.iter(qn('w:t')))
            for match in reversed(list(re.finditer(r'(?<!\w)(?:X,\s*XXX|X{1,4})(?=\s+square feet)', text, re.I))):
                for node in replace_span(p._p, match.start(), match.end(), min_lot):
                    _blacken_run(Run(node.getparent(), p))
        if 'Planner' in t or '393-' in t:
            _replace_runs_contextual(p, exact_map={'Planner’s Name': ctx['Planner_Name'], 'Planner Name': ctx['Planner_Name'],
                                                   "Planner's Name": ctx['Planner_Name']},
                                     contains_map={'Planner’s Name': ctx['Planner_Name'], 'Planner Name': ctx['Planner_Name'],
                                                   'or call the ': 'or call '})
            _fill_phone(p, ctx['Planner_Phone'])
        _replace_in_runs(p, common)
    _replace_text_nodes(doc, {k: ctx['Planner_Email'] for k in _EMAIL_PLACEHOLDERS})
    _fix_mailto_links(doc, ctx['Planner_Email'])


def generate_owner_letters(excel_path, template_path, map_path, output_path, preview=False) -> int:
    """Mail merge for the notification, protest or public-hearing letter (the
    template decides which).  One section block per owner, map after each letter."""
    ctx = build_context(excel_path)
    _ci, owners = ME.load_case_and_owners(excel_path)
    if not owners:
        raise ValueError('No valid owner rows found on Step 1.')
    if not map_path or not Path(map_path).exists():
        raise ValueError('Select a boundary map in the wizard before generating letters.')
    selected = owners[:1] if preview else owners
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        map_png = _map_to_png(map_path, td)
        base = None
        for owner in selected:
            doc = Document(str(template_path))
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
    return len(selected)


# =============================================================================
# Applicant sign letter
# =============================================================================
def generate_applicant_letter(excel_path, template_path, output_path) -> int:
    ctx = build_context(excel_path)
    doc = Document(str(template_path))
    for p in doc.paragraphs:
        t = _norm(p.text)
        ts = t.strip()
        if ts == 'Date':
            _set_exact_para_preserve_first_run(p, ctx['Letter_Date'])
        elif ts == 'Name':
            _set_exact_para_preserve_first_run(p, ctx['Applicant_Name'])
        elif ts == 'Address':
            _set_exact_para_preserve_first_run(p, ctx['Applicant_Address_Block'])
        elif 'Houston, TX XXXX' in t:
            _replace_in_runs(p, {'Houston, TX XXXX': ctx['Applicant_City_State_Zip']})
        elif 'block of XX Street' in t:
            _replace_in_runs(p, {'XX block of XX Street, X and X sides,': ctx['Area_Line_1']})
        elif ts.startswith('between X Street'):
            _replace_in_runs(p, {'between X Street and X Drive': ctx['Area_Line_2']})
        elif ts.startswith('Dear'):
            _replace_runs_contextual(p, exact_map={'Property Owner': ctx['Applicant_Name']})
        elif 'minimum lot size is' in t:
            _replace_runs_contextual(p, exact_map={
                'xx': ctx['Minimum_Lot_Size_Text'], 'XX': ctx['Minimum_Lot_Size_Text'],
                'two (2': f"{ctx['Sign_Count_Words']} ({ctx['Sign_Count_Number']}",
                'one': ctx['Sign_Placement']},
                contains_map={' sign at each corner.': '.'})
        elif 'feel free to contact' in t:
            _replace_runs_contextual(p, exact_map={'Planner Name': ctx['Planner_Name'], 'XXXX': '', 'xxxx': ''},
                                     contains_map={' at (832) 393-': f" at {ctx['Planner_Phone']}"})
    _replace_text_nodes(doc, {k: ctx['Planner_Email'] for k in _EMAIL_PLACEHOLDERS})
    _fix_mailto_links(doc, ctx['Planner_Email'])
    _clear_all_red_run_colors(doc)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    unlock_docx(out)
    return 1


# =============================================================================
# Director approval (2026 form)
# =============================================================================
def generate_director_approval(excel_path, template_path, output_path) -> int:
    ctx = build_context(excel_path)
    from core.city_templates import render_updated
    if render_updated(template_path, output_path, ctx, excel_path, 'mls', 'director'): return 1
    doc = Document(str(template_path))
    body_ids = {id(p._p) for p in doc.paragraphs}
    for p in _iter_all_paragraphs(doc):
        t = _norm(p.text)
        ts = t.strip()
        if 'Lot Size Block No.' in t:
            _replace_runs_contextual(p, exact_map={'XXX': ctx['Application_Number']})
        elif id(p._p) in body_ids and ' block of ' in t and 'between' in t:
            _set_exact_para_preserve_first_run(p, ctx['Area_Description_Period'])
        elif 'minimum lot size of' in t:
            _replace_runs_contextual(p, exact_map={'XXXX': ctx['Minimum_Lot_Size_Text']})
        elif 'Prepared by' in t:
            _replace_runs_contextual(p, exact_map={'Planner Name': ctx['Planner_Name']})
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


# =============================================================================
# Routing form (2026 coversheet routing form)
# =============================================================================
def _fill_sdts_in_order(p, values):
    """Assign the non-checkbox content controls of one paragraph, in order."""
    sdts = [s for s in p._p.iter(qn('w:sdt')) if not _sdt_is_checkbox(s)
            and not any(a.tag == qn('w:sdt') for a in s.iterancestors() if a is not p._p)]
    for sdt, value in zip(sdts, values):
        if value is None:
            continue
        first_run = sdt.find(f'.//{qn("w:r")}')
        new_run = OxmlElement('w:r')
        if first_run is not None and first_run.find(qn('w:rPr')) is not None:
            rpr = copy.deepcopy(first_run.find(qn('w:rPr')))
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


def generate_routing_form(excel_path, template_path, output_path) -> int:
    ctx = build_context(excel_path)
    from core.city_templates import render_updated
    if render_updated(template_path, output_path, ctx, excel_path, 'mls', 'routing'): return 1
    doc = Document(str(template_path))
    is_pc = ctx['Approval_Path'] == 'Planning Commission Approval'
    location = f"The {ctx['Area_Description']}" if ctx.get('Block_Description') else 'The X block of X Street, X and X sides, between X Street and X Street'
    try:
        sfr_ok = float(str(ctx['SFR_Percentage']).replace('%', '')) >= 60
    except Exception:
        sfr_ok = True
    for i, p in enumerate(doc.paragraphs):
        t = _norm(p.text)
        if t.startswith('The topic of Request'):
            _fill_sdts_in_order(p, [f"Special Minimum Lot Size Block Application No. {ctx['Application_Number']}"])
            if i + 1 < len(doc.paragraphs):
                loc = doc.paragraphs[i + 1]
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
            _set_paragraph_checkboxes(p, [False, True])
            _fill_sdts_in_order(p, [ctx['Application_Number'], f": {ctx['Minimum_Lot_Size_Text']}",
                                    ctx['Council_District'] or 'XXX'])
        elif 'at least 60 percent' in t:
            _set_paragraph_checkboxes(p, [sfr_ok, not sfr_ok])
        elif 'SMLS Area Application' in t:
            _set_paragraph_checkboxes(p, [False, True, False])
        elif 'SMLS Block Renewal' in t and 'SMBL' not in t:
            _set_paragraph_checkboxes(p, [False])
        elif 'SMBL Block Renewal' in t:
            _set_paragraph_checkboxes(p, [False])
        elif 'Planning Commission Approval' in t and 'Director' in t:
            _set_paragraph_checkboxes(p, [is_pc, not is_pc, True, True, False])
    _clear_all_red_run_colors(doc)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    prepare_routing_form(doc)
    doc.save(str(out))
    unlock_docx(out)
    return 1


# =============================================================================
# OneMeeting staff report (Director path)
# =============================================================================
def generate_staff_report(excel_path, template_path, output_path) -> int:
    ctx = build_context(excel_path)
    from core.city_templates import render_updated
    if render_updated(template_path, output_path, ctx, excel_path, 'mls', 'staff'): return 1
    doc = Document(str(template_path))
    area = ctx['Area_Description']
    caption_area = ctx['Ordinance_Area_Lower']
    district = ctx['Council_District'] or 'X'
    subdivision = ctx['Subdivision_Name']
    if subdivision != 'SUBDIVISION NAME' and 'subdivision' not in subdivision.lower():
        subdivision = f"{subdivision} Subdivision"
    for p in doc.paragraphs:
        t = _norm(p.text)
        c = _RunCursor(p)
        if t.startswith('DISTRICT:') and ctx['Council_District']:
            c.set('CHOOSE THE D', ctx['Council_District'])
            c.set('ISTRICT', '')
        elif 'Block Application No.' in t:
            _replace_runs_contextual(p, exact_map={'XXX': ctx['Application_Number']})
        elif t.startswith('(') and 'block of' in t:
            c.set('2800 - 2900 block of Wheeler Street, north side, between Delano Street and Ennis Street', area)
        elif t.startswith('ITEM CAPTION'):
            c.set('ENTER THE CAPTION FROM THE ORDINANCE (see below example)', '')
            c.set('the north side of the 2800-2900 Blocks of Wheeler Street, between Delano Street and Ennis Street ', f"the {caption_area} ")
            c.set('DISTRICT D - EVANS-SHABAZZ', f"DISTRICT {district} - {ctx['Council_Member']}")
        elif t.startswith('In accordance with'):
            c.set('XXXX XXX', ctx['Applicant_Property_Address'])
            c.set(' ', '')
            c.set('Street, of the ', ', of the ')
            c.set('SUBDIVISION NAME ', f"{subdivision} ")
            c.span('XXXX', 'Street ', area + ' ')
        elif t.startswith('As no action is required'):
            c.set('XXXX', ctx['Minimum_Lot_Size_Text'])
            c.span('XXXX', 'Street', area)
        elif 'Your Name, Position, Phone' in t:
            _replace_in_runs(p, {'Your Name, Position, Phone Number': f"{ctx['Planner_Name']}, {ctx['Planner_Position']}, {ctx['Planner_Phone']}"})
    _clear_all_red_run_colors(doc)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    unlock_docx(out)
    return 1


# =============================================================================
# Draft ordinance (2026 template with content controls)
# =============================================================================
def _sdt_text(sdt):
    return _norm(''.join(t.text or '' for t in sdt.iter(qn('w:t')))).strip()


def _unwrap_block_sdt(sdt):
    """Replace a block-level content control by its content children."""
    parent = sdt.getparent()
    content = sdt.find(qn('w:sdtContent'))
    idx = parent.index(sdt)
    kids = list(content) if content is not None else []
    for k in kids:
        parent.insert(idx, k)
        idx += 1
    parent.remove(sdt)
    return kids


def _fill_ordinance_exhibits(doc, ctx, rows, map_path):
    body = doc.element.body
    for sdt in [s for s in body if s.tag == qn('w:sdt')]:
        text = _sdt_text(sdt)
        tables = list(sdt.iter(qn('w:tbl')))
        drawings = list(sdt.iter(qn('w:drawing')))
        if tables:
            # Exhibit C property list: ADDRESS | HCAD | LOT/TRACT | BLOCK | SUBDIVISION
            kids = _unwrap_block_sdt(sdt)
            tbl = next(k for k in kids if k.tag == qn('w:tbl'))
            trs = tbl.findall(qn('w:tr'))
            proto = copy.deepcopy(trs[1]) if len(trs) > 1 else copy.deepcopy(trs[0])
            for tr in trs[1:]:
                tbl.remove(tr)
            for row in rows[:300]:
                tr = copy.deepcopy(proto)
                vals = [row.get('ADDRESS', ''), row.get('HCAD', ''), row.get('LOT_TRACT', ''), row.get('BLOCK', ''), row.get('SUBDIVISION', '')]
                for tc, val in zip(tr.findall(qn('w:tc')), vals):
                    ps = tc.findall(qn('w:p'))
                    for extra in ps[1:]:
                        tc.remove(extra)
                    p = Paragraph(ps[0], doc)
                    if p.runs:
                        _set_run_text_preserve_xml(p.runs[0], val)
                        for r in p.runs[1:]:
                            _set_run_text_preserve_xml(r, '')
                    else:
                        p.add_run(val)
                tbl.append(tr)
        elif drawings:
            # Exhibit B boundary map
            kids = _unwrap_block_sdt(sdt)
            if map_path and Path(map_path).exists():
                with tempfile.TemporaryDirectory() as td:
                    png = _map_to_png(map_path, td)
                    p_el = next((k for k in kids if k.tag == qn('w:p')), None)
                    if png and p_el is not None:
                        p = Paragraph(p_el, doc)
                        p.clear()
                        p.alignment = 1
                        p.add_run().add_picture(png, width=Inches(6.25))
        elif 'enter text' in text.lower() or not text:
            # Exhibit A property descriptions
            kids = _unwrap_block_sdt(sdt)
            anchor = next((k for k in kids if k.tag == qn('w:p')), None)
            if anchor is None:
                continue
            descs = []
            for row in rows:
                d = row.get('LEGAL_DESCRIPTION', '')
                if d and d not in descs:
                    descs.append(d)
            first = Paragraph(anchor, doc)
            if descs:
                if first.runs:
                    _set_run_text_preserve_xml(first.runs[0], descs[0])
                    for r in first.runs[1:]:
                        _set_run_text_preserve_xml(r, '')
                else:
                    first.add_run(descs[0])
                _blacken_run(first.runs[0])
                prev = anchor
                for d in descs[1:]:
                    new_p = copy.deepcopy(anchor)
                    prev.addnext(new_p)
                    np_ = Paragraph(new_p, doc)
                    _set_run_text_preserve_xml(np_.runs[0], d)
                    prev = new_p


def generate_draft_ordinance(excel_path, template_path, map_path, output_path) -> int:
    ctx = build_context(excel_path)
    doc = Document(str(template_path))
    rows = ME._read_property_rows_for_ordinance(excel_path)
    for row in rows:
        digits = re.sub(r'\D+', '', row.get('HCAD', '') or '')
        if digits:
            row['HCAD'] = digits.zfill(13)  # HCAD account numbers carry a leading zero
    # caption (tagged control)
    for sdt in doc.element.body.iter(qn('w:sdt')):
        tag = sdt.find('.//' + qn('w:tag'))
        if tag is not None and tag.get(qn('w:val')) == 'MLS Boundary Description':
            texts = [t for t in sdt.iter(qn('w:t'))]
            if texts:
                texts[0].text = ctx['Ordinance_Area_Upper']
                for extra in texts[1:]:
                    extra.text = ''
            for color_el in sdt.iter(qn('w:color')):
                _blacken_color_el(color_el)
    lot_phrase = ctx['Lot_Size_Phrase']
    rules = [
        ('51%', f"{ctx['Support_Percentage']}%", None),
        ('sixteen (16) out of twenty-five (25)', ctx['Qualifying_Phrase'], None),
        ('six thousand eight hundred and twenty (6,820)', lot_phrase, None),
        ('6,820 feet.', f"{ctx['Minimum_Lot_Size_Text']} square feet.", None),
    ]
    if _clean(ctx.get('Legal_File_No')):
        rules.append(('XXXXXXXXX', _clean(ctx.get('Legal_File_No')), 'L.D. File'))
    # the lower-case area control has sample wording; match it by paragraph context
    for sdt in list(doc.element.body.iter(qn('w:sdt'))):
        if _sdt_is_checkbox(sdt):
            continue
        txt = _sdt_text(sdt)
        if txt.lower().startswith('north and south sides of the 4900'):
            rules.append((txt, ctx['Ordinance_Area_Lower'], None))
    _fill_content_controls(doc, rules)
    for p in doc.paragraphs:
        t = _norm(p.text)
        if t.startswith('Requested by'):
            _replace_in_runs(p, {'Vonn Tran': ctx['Director_Name']})
        elif 'Ordinance No.' in t:
            if _clean(ctx.get('Ordinance_Year')):
                _replace_in_runs(p, {'2026': _clean(ctx.get('Ordinance_Year'))})
            # the template paragraph carries a stray automatic list number and a
            # hidden SEQ CHAPTER field that renders as a trailing "1"
            pPr = p._p.get_or_add_pPr()
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                pPr.remove(numPr)
            for r in list(p._p.findall(qn('w:r'))):
                if r.find(qn('w:fldChar')) is not None or r.find(qn('w:instrText')) is not None:
                    p._p.remove(r)
    # content-control text that keeps the template's red placeholder colour
    for sdt in doc.element.body.iter(qn('w:sdt')):
        if not _sdt_is_checkbox(sdt):
            for color_el in sdt.iter(qn('w:color')):
                _blacken_color_el(color_el)
    _fill_ordinance_exhibits(doc, ctx, rows, map_path)
    _clear_all_red_run_colors(doc)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    unlock_docx(out)
    return 1


# =============================================================================
# RCA packet (2026 order)
# =============================================================================
def generate_rca_packet_pdf(excel_path, templates, map_path, tax_map_path, output_pdf,
                            subdivision_plat_path='', hcad_records=None, include_staff_analysis=False) -> int:
    """Routing form, Director or PC approval, boundary map, tax map, then any
    optional extras (plat, Step 2 staff analysis, HCAD records)."""
    ctx = build_context(excel_path)
    output_pdf = Path(output_pdf)
    from core.packet_work import packet_workspace
    with packet_workspace() as work:
        if not str(map_path or '').strip():
            raise ValueError('Select the boundary map before building the RCA packet.')
        if not str(tax_map_path or '').strip():
            raise ValueError('Select the tax map before building the RCA packet.')
        routing = work / '01_MLS_Routing_Form.docx'
        generate_routing_form(excel_path, templates['routing'], routing)
        pdfs = [routing_pdf_one_page(ME.ensure_pdf(routing, work), work)]
        if ctx['Approval_Path'] == 'Planning Commission Approval':
            pc = work / '02_Planning_Commission_Approval.docx'
            ME.generate_planning_commission_approval(excel_path, templates['pc'], pc)
            pdfs.append(ME.ensure_pdf(pc, work))
        else:
            da = work / '02_Director_Approval.docx'
            generate_director_approval(excel_path, templates['director'], da)
            pdfs.append(ME.ensure_pdf(da, work))
        pdfs.append(ME.ensure_pdf(map_path, work))
        pdfs.append(ME.ensure_pdf(tax_map_path, work))
        if str(subdivision_plat_path or '').strip():
            pdfs.append(ME.ensure_pdf(subdivision_plat_path, work))
        if include_staff_analysis:
            staff = work / '05_Staff_Analysis_Step2.pdf'
            ME.export_step2_staff_analysis_pdf(excel_path, staff)
            pdfs.append(staff)
        hcad = [p for p in (hcad_records or []) if str(p).strip()]
        if hcad:
            pdfs.extend(condense_hcad_pdfs([ME.ensure_pdf(p, work) for p in hcad], work))
        ME.merge_pdfs(pdfs, output_pdf)
        return len(pdfs)
