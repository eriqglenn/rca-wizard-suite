
from pathlib import Path
import tempfile, subprocess, shutil, sys, datetime
from copy import deepcopy
import openpyxl
from core.routing_form import prepare_routing_form, routing_pdf_one_page
from docx import Document
from core.docx_unlock import unlock_docx
from core.pdf_condense import condense_hcad_pdfs
from core.sheet_layout import apply_staff_analysis_page_setup
from docx.shared import Inches, RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import fitz

def _format_date(v):
    """Format an Excel date/datetime as the City's long form, e.g.
    'January 18, 2006'. Built from parts (no %-d/%#d) so it is identical on
    Windows and Linux, and so a datetime never prints as
    '2006-01-18 00:00:00' in letters or forms."""
    return f"{v.strftime('%B')} {v.day}, {v.year}"

def clean(v):
    if v is None: return ''
    if isinstance(v, datetime.datetime): return _format_date(v)
    if isinstance(v, datetime.date): return _format_date(v)
    if isinstance(v,float) and v.is_integer(): return str(int(v))
    return str(v).strip()

def pdf_first_page_to_png(pdf_path, png_path):
    doc=fitz.open(str(pdf_path)); page=doc[0]
    pix=page.get_pixmap(matrix=fitz.Matrix(2,2), alpha=False)
    pix.save(str(png_path)); doc.close()
    return png_path

def map_image_for_docx(map_path, temp_dir):
    """Return an image path suitable for python-docx insertion.
    PDF maps are converted into the OS temp folder only; no PNG is left in the user's output folder.
    """
    mp = Path(map_path)
    if str(mp).lower().endswith('.pdf'):
        tmp = Path(temp_dir) / f"{mp.stem}_page1.png"
        pdf_first_page_to_png(mp, tmp)
        return str(tmp)
    return str(mp)

_EXCEL_ERRORS = {'#NAME?', '#VALUE!', '#REF!', '#N/A', '#DIV/0!', '#NULL!', '#NUM!'}


def _read_step2_caseinfo_overlay(wb_path):
    """Read the official auto-populate values straight from Step 2.

    Step 2 is the single source of truth. The MBL Case Info tab's "Auto-Read From
    Step 2" cells are live Excel formulas that reference Step 2, but openpyxl reads
    only Excel's cached formula results, which can be stale or show #NAME? if the
    workbook was edited and not recalculated. To guarantee Step 2 always feeds the
    forms (and that Step 2 and Case Info can never disagree), the engine reads the
    calculated values directly from Step 2 here and overlays them onto Case Info.

    Planner-entered narrative/contact fields (dates, planner info, applicant info,
    approval path, findings) are NOT overlaid; those remain owned by Case Info.
    """
    out = {}
    try:
        wb = openpyxl.load_workbook(wb_path, data_only=True)
    except Exception:
        return out
    if 'Step 2 - Main Calculation Sheet' not in wb.sheetnames:
        return out
    ws = wb['Step 2 - Main Calculation Sheet']
    s2 = read_step2(wb_path)

    def put(key, value):
        v = clean(value)
        if v and v.upper() not in _EXCEL_ERRORS:
            out[key] = v

    put('Application_Number', ws['B2'].value)
    put('Application_Type', 'Special Minimum Building Line Block Application')
    put('Street_Name', ws['B4'].value)
    put('Block_Description', s2.get('Block_Description', ''))
    put('Side_Of_Street', s2.get('Side_Of_Street', ''))
    put('Cross_Streets', s2.get('Cross_Streets', ''))
    put('Minimum_Building_Line_Feet', s2.get('Minimum_Building_Line_Feet', ''))
    return out


def read_case_info(wb_path):
    wb=openpyxl.load_workbook(wb_path, data_only=True)
    ci={}
    if 'Case Info' in wb.sheetnames:
        ws=wb['Case Info']
        for r in range(1, ws.max_row+1):
            k=ws.cell(r,1).value; v=ws.cell(r,2).value
            if k and str(k).strip() not in ['Field','Purpose','Planner Inputs','Applicant / Sign Posting','Approvals Narrative','MBL Wizard — Case Info']:
                ci[str(k).strip()] = clean(v)
    # Step 2 is the source of truth for calculated/auto-read fields. Overlay the
    # live Step 2 values last so a stale Case Info formula cache cannot win.
    for k, v in _read_step2_caseinfo_overlay(wb_path).items():
        ci[k] = v
    return ci

def header_map(ws):
    return {str(ws.cell(1,c).value).strip(): c for c in range(1, ws.max_column+1) if ws.cell(1,c).value is not None}

def read_owner_rows(wb_path):
    wb=openpyxl.load_workbook(wb_path, data_only=True)
    ws=wb['Step 1 - Input Properties Info']
    hm=header_map(ws); rows=[]
    for r in range(2, ws.max_row+1):
        owner=clean(ws.cell(r, hm.get('OWNER_MAILTO')).value) if 'OWNER_MAILTO' in hm else ''
        mail1=clean(ws.cell(r, hm.get('Mail_Addr_1')).value) if 'Mail_Addr_1' in hm else ''
        if not owner or not mail1 or owner=='0': continue
        rows.append({'Mailto': owner,'Mail_Addr_1': mail1,'Mail_Addr_2': clean(ws.cell(r, hm.get('Mail_Addr_2')).value) if 'Mail_Addr_2' in hm else '', 'Mail_City': clean(ws.cell(r, hm.get('Mail_City')).value) if 'Mail_City' in hm else '', 'Mail_State': clean(ws.cell(r, hm.get('Mail_State')).value) if 'Mail_State' in hm else '', 'Mail_Zip': clean(ws.cell(r, hm.get('Mail_Zip')).value) if 'Mail_Zip' in hm else ''})
    return rows

def read_step2(wb_path):
    wb=openpyxl.load_workbook(wb_path, data_only=True)
    ws=wb['Step 2 - Main Calculation Sheet']
    block=clean(ws['D4'].value); cross1=clean(ws['B5'].value); cross2=clean(ws['D5'].value); side1=clean(ws['B6'].value); side2=clean(ws['D6'].value)
    sides=side1 if not side2 else f"{side1} and {side2} sides"
    sides=sides.replace('  ',' ').strip()
    if not sides.lower().endswith('side') and not sides.lower().endswith('sides'): sides+=' side'
    min_bl=clean(ws['L16'].value) or clean(ws['E4'].value).replace('-Foot','').replace('Foot','')
    try: min_bl=str(int(float(min_bl)))
    except Exception: pass
    cross=f"between {cross1} and {cross2}" if cross1 or cross2 else ''
    return {'Block_Description': block,'Side_Of_Street': sides,'Cross_Streets': cross,'Minimum_Building_Line_Feet': min_bl,'Subject_Line_1': 'Special Minimum Building Line Block Application','Subject_Line_2': f"{block}, {sides},".replace(' ,',','),'Subject_Line_3': cross}

def build_address(owner):
    lines=[owner['Mailto'],owner['Mail_Addr_1']]
    if owner.get('Mail_Addr_2') and owner.get('Mail_Addr_2') not in ['0','None']: lines.append(owner['Mail_Addr_2'])
    lines.append(f"{owner.get('Mail_City','')}, {owner.get('Mail_State','')} {owner.get('Mail_Zip','')}")
    return '\n'.join([l.strip() for l in lines if l and l.strip()])

def remove_trailing_blank_paragraphs(doc):
    body=doc.element.body
    for child in reversed(list(body)):
        if child.tag.endswith('sectPr'):
            continue
        if child.tag.endswith('p'):
            texts=''.join(t.text or '' for t in child.iter() if t.tag.endswith('}t'))
            drawings=[x for x in child.iter() if x.tag.endswith('}drawing')]
            brs=[x for x in child.iter() if x.tag.endswith('}br')]
            if not texts.strip() and not drawings and not brs:
                body.remove(child); continue
        break

def set_run_font(run, size=12):
    # Force the generated body text to Arial 12 so inserted values match the MBL template expectation.
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    try:
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.rFonts
        if rFonts is None:
            from docx.oxml import OxmlElement
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', 'Arial')
        rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi', 'Arial')
    except Exception:
        pass

def set_para_text(p,text,color=(0,0,0), align_left=False):
    p._p.clear_content()
    if align_left:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        # 0.5" matches the template's own firstLine=720 twips for the address
        # block, so all address lines sit exactly where the template puts the
        # first line (was 0.55", which drifted 0.05" right of the template).
        p.paragraph_format.left_indent = Inches(0.5)
        p.paragraph_format.first_line_indent = Inches(0)
    run=p.add_run()
    run.text=text
    set_run_font(run, 12)
    if color: run.font.color.rgb=RGBColor(*color)

def normalize_generated_text(doc):
    # Keep generated body text consistent as Arial 12. Do not touch header/footer branding.
    container = doc
    for p in container.paragraphs:
        for r in p.runs:
            set_run_font(r, 12)
    for t in container.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        set_run_font(r, 12)

def _clear_numbering(p):
    try:
        pPr = p._p.get_or_add_pPr()
        numPr = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr')
        if numPr is not None:
            pPr.remove(numPr)
    except Exception:
        pass

def _materialize_next_steps_numbering(doc):
    # The original template uses Word automatic numbering. When letters are appended into
    # one combined document, Word continues the numbering across owner letters. Replace
    # those two list items with fixed text so every letter reads 1) and 2).
    targets = [
        ('Confirm your support or objection:', '1)  '),
        ('Watch for Further Notices:', '2)  '),
    ]
    for p in doc.paragraphs:
        txt = p.text.strip()
        for marker, prefix in targets:
            if txt.startswith(marker):
                _clear_numbering(p)
                set_para_text(p, prefix + txt, align_left=True)
                break

def replace_text_in_doc(doc, replacements):
    def do_para(p):
        txt=p.text; new=txt
        for old,val in replacements.items():
            if old in new: new=new.replace(old,val)
        if new!=txt:
            align_address = ('ADDRESS BLOCK' in txt)
            set_para_text(p,new, align_left=align_address)
    for p in doc.paragraphs: do_para(p)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs: do_para(p)
    _materialize_next_steps_numbering(doc)
    # Do NOT normalize the full document; it changes the City of Houston letterhead/logo text.
    # Formatting is applied only to the paragraphs where replacement text is inserted.

def fill_letter(template_path, wb_path, map_path, owner, out_path):
    ci=read_case_info(wb_path); s2=read_step2(wb_path)
    doc=Document(template_path)
    replacements={'ADDRESS BLOCK': build_address(owner),'Date': ci.get('Letter_Date',''),'#### blocks of X Street, south side,': s2['Subject_Line_2'],'between Y Street and Z Street': s2['Subject_Line_3'],'X feet': f"{s2['Minimum_Building_Line_Feet'] or 'X'} feet",'Month XX, 2026': ci.get('Protest_Deadline',''),'Planner Name@houstontx.gov': ci.get('Planner_Email',''),'Planner Name': ci.get('Planner_Name',''),'832-393-XXXX': ci.get('Planner_Phone','')}
    replace_text_in_doc(doc,replacements)
    with tempfile.TemporaryDirectory() as td:
        for p in doc.paragraphs:
            if 'Insert Map' in p.text:
                set_para_text(p,'', align_left=False)
                run=p.add_run(); mp=map_image_for_docx(map_path, td)
                run.add_picture(mp, width=Inches(5.9))
        remove_trailing_blank_paragraphs(doc); doc.save(out_path)
    unlock_docx(out_path)  # fully editable output
    return out_path

_WT = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _render_single_owner_letter(template_path, wb_path, owner, embedded_map):
    """Render exactly one MBL owner letter from a FRESH copy of the official template.

    This is the same code path used for the Preview One Letter output, so every
    letter keeps the template's exact section structure and its first-page
    council-member footer. Nothing here edits margins, page size, section
    properties, footer references, or the City letterhead/logo.
    """
    ci = read_case_info(wb_path)
    s2 = read_step2(wb_path)
    doc = Document(template_path)
    replacements = {
        'ADDRESS BLOCK': build_address(owner),
        'Date': ci.get('Letter_Date', ''),
        '#### blocks of X Street, south side,': s2['Subject_Line_2'],
        'between Y Street and Z Street': s2['Subject_Line_3'],
        'X feet': f"{s2['Minimum_Building_Line_Feet'] or 'X'} feet",
        'Month XX, 2026': ci.get('Protest_Deadline', ''),
        'Planner Name@houstontx.gov': ci.get('Planner_Email', ''),
        'Planner Name': ci.get('Planner_Name', ''),
        '832-393-XXXX': ci.get('Planner_Phone', ''),
    }
    replace_text_in_doc(doc, replacements)
    for p in doc.paragraphs:
        if 'Insert Map' in p.text:
            set_para_text(p, '', align_left=False)
            run = p.add_run()
            run.add_picture(embedded_map, width=Inches(5.9))
    remove_trailing_blank_paragraphs(doc)
    return doc


def _final_sectpr(body):
    last = body[-1] if len(body) else None
    return last if last is not None and last.tag == _WT + 'sectPr' else None


def _append_letter_as_own_section(master, letter):
    """Append one rendered MBL letter to master as its own self-contained section.

    The MBL owner-notification template (MBL_ALL_OWNERS.docx) stores the current
    council-member footer as a first-page footer (w:titlePg + footerReference
    type="first") on its single section. A true mail merge must give EVERY owner
    letter its own copy of that section structure so the footer (and the template
    body margins) repeat on every letter -- not just the first. The previous
    combiner pasted later letters as bare body paragraphs after a manual page
    break, collapsing everything to one section; the current council-member
    footer then showed only on owner 1 and later letters lost it. This routine
    instead preserves each letter's own sectPr block (and thus its footer
    references) verbatim.

    All letters in one case share the identical council-member footer, so they
    correctly reference the single template footer part already present in the
    master document.
    """
    mbody = master.element.body
    lbody = letter.element.body

    # Close the master's current trailing content into its own section by tucking
    # the master's existing final sectPr into a new section-break paragraph.
    msect = _final_sectpr(mbody)
    sect_break_p = master.element.makeelement(_WT + 'p', {})
    ppr = master.element.makeelement(_WT + 'pPr', {})
    sect_break_p.append(ppr)
    if msect is not None:
        mbody.remove(msect)
        ppr.append(msect)
    mbody.append(sect_break_p)

    # Append the next letter's body content, preserving any intermediate sectPr,
    # so its margins and first-page footer repeat. The letter's own final sectPr
    # becomes the master's new final sectPr.
    lsect = _final_sectpr(lbody)
    for child in list(lbody):
        if child is lsect:
            continue
        mbody.append(deepcopy(child))
    if lsect is not None:
        mbody.append(deepcopy(lsect))


def generate_all(template_path, wb_path, map_path, out_path):
    owners = read_owner_rows(wb_path)
    if not owners:
        raise RuntimeError('No owner records found.')
    out_path = Path(out_path)
    with tempfile.TemporaryDirectory() as td:
        # Convert a PDF map once into the OS temp folder. It is embedded in the
        # DOCX; the temporary PNG is removed automatically when this function ends.
        embedded_map = map_image_for_docx(map_path, td)

        # True mail merge. Owner 1 is rendered from a fresh template copy (keeping
        # its native footer relationships). Each subsequent owner is rendered the
        # same way and appended as its OWN section block, so every letter repeats
        # the template body margins AND its first-page council-member footer.
        master = _render_single_owner_letter(template_path, wb_path, owners[0], embedded_map)
        for owner in owners[1:]:
            letter = _render_single_owner_letter(template_path, wb_path, owner, embedded_map)
            _append_letter_as_own_section(master, letter)
        master.save(str(out_path))
    unlock_docx(out_path)  # fully editable output
    return out_path, len(owners)


def _replace_all_text_preserve_template(doc, replacements):
    # Replace placeholders in paragraphs/tables. Only paragraphs that receive replacements
    # are rewritten; headers/letterhead and unrelated template formatting stay untouched.
    def do_para(p):
        txt = p.text
        new = txt
        for old, val in replacements.items():
            if old in new:
                new = new.replace(old, str(val))
        if new != txt:
            p._p.clear_content()
            run = p.add_run(new)
            set_run_font(run, 12)
            run.font.color.rgb = RGBColor(0,0,0)
    for p in doc.paragraphs:
        do_para(p)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    do_para(p)

def _applicant_address_lines(ci):
    lines = [ci.get('Applicant_Name',''), ci.get('Applicant_Address_1','')]
    addr2 = ci.get('Applicant_Address_2','')
    if addr2 and addr2 not in ['0','None','none','NONE']:
        lines.append(addr2)
    city_line = f"{ci.get('Applicant_City','')}, {ci.get('Applicant_State','')} {ci.get('Applicant_Zip','')}".strip()
    lines.append(city_line)
    return [l.strip() for l in lines if l and l.strip() and l.strip() not in ['None','0']]


def generate_applicant_letter(template_path, wb_path, out_path):
    ci = read_case_info(wb_path)
    s2 = read_step2(wb_path)
    doc = Document(template_path)
    addr_lines = _applicant_address_lines(ci)
    applicant_name = addr_lines[0] if len(addr_lines)>0 else ci.get('Applicant_Name','')
    addr1 = addr_lines[1] if len(addr_lines)>1 else ci.get('Applicant_Address_1','')
    addr2 = ci.get('Applicant_Address_2','')
    if addr2 and addr2 not in ['0','None','none','NONE']:
        addr_par = addr1 + '\n' + addr2
        city_line = f"{ci.get('Applicant_City','')}, {ci.get('Applicant_State','')} {ci.get('Applicant_Zip','')}".strip()
    else:
        addr_par = addr1
        city_line = f"{ci.get('Applicant_City','')}, {ci.get('Applicant_State','')} {ci.get('Applicant_Zip','')}".strip()
    exact = {
        'Date': ci.get('Letter_Date',''),
        'Name': applicant_name,
        'Address': addr_par,
        'Houston, TX XXXX': city_line,
    }
    inline = {
        'XX block of X Street, X and X side(s),': f"{s2.get('Block_Description','')}, {s2.get('Side_Of_Street','')},",
        'between XX and XX': s2.get('Cross_Streets',''),
        'Property Owners Name': applicant_name,
        'X feet': f"{s2.get('Minimum_Building_Line_Feet','') or 'X'} feet",
        'two (2)': f"{ci.get('Sign_Count_Words','two')} ({ci.get('Sign_Count_Number','2')})",
        'one sign at each corner': ci.get('Sign_Placement','one sign at each corner'),
        'Planner Name@houstontx.gov': ci.get('Planner_Email',''),
        '(832) 393-XXXX': ci.get('Planner_Phone',''),
        '832-393-XXXX': ci.get('Planner_Phone',''),
        'Planner at': ci.get('Planner_Name','') + ' at',
    }
    def set_text(p, text):
        p._p.clear_content()
        run = p.add_run(str(text))
        set_run_font(run, 12)
        run.font.color.rgb = RGBColor(0,0,0)
    def do_para(p):
        stripped=p.text.strip()
        if stripped in exact:
            set_text(p, exact[stripped]); return
        txt=p.text; new=txt
        for old,val in inline.items():
            if old in new:
                new=new.replace(old, str(val))
        if new != txt:
            set_text(p,new)
    for p in doc.paragraphs:
        do_para(p)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    do_para(p)
    remove_trailing_blank_paragraphs(doc)
    doc.save(out_path)
    unlock_docx(out_path)  # fully editable output
    return out_path


# ---- MBL approval generation ----
def _num_words(n):
    try: n=int(float(n))
    except Exception: return str(n)
    words={0:'zero',1:'one',2:'two',3:'three',4:'four',5:'five',6:'six',7:'seven',8:'eight',9:'nine',10:'ten',11:'eleven',12:'twelve',13:'thirteen',14:'fourteen',15:'fifteen',16:'sixteen',17:'seventeen',18:'eighteen',19:'nineteen',20:'twenty',21:'twenty-one',22:'twenty-two',23:'twenty-three',24:'twenty-four',25:'twenty-five',26:'twenty-six',27:'twenty-seven',28:'twenty-eight',29:'twenty-nine',30:'thirty'}
    if n in words:
        return words[n]
    if 0 <= n < 1000000:
        from wizards.mls.mailout_engine import number_to_words
        return number_to_words(n)
    return str(n)

def _pct_str(v):
    try:
        f=float(v)
        if f <= 1: f *= 100
        return str(int(round(f)))
    except Exception:
        s=str(v).strip().replace('%','')
        return s if s else '0'

def read_step2_approval(wb_path):
    wb = openpyxl.load_workbook(wb_path, data_only=True)
    ws = wb['Step 2 - Main Calculation Sheet']
    basic = read_step2(wb_path)
    app_no = clean(ws['B2'].value) or 'XX'
    min_bl = basic.get('Minimum_Building_Line_Feet','')
    # support percentage: use visible Step 2 F8 if cached; otherwise compute from signed support and lot area.
    support_val = ws['F8'].value
    if not support_val:
        signed_area = 0; total_area = 0
        for r in range(17, 101):
            area = ws.cell(r,4).value
            try: area = float(area or 0)
            except Exception: area = 0
            if area > 0: total_area += area
            signed = str(ws.cell(r,3).value or '').strip().upper()
            if signed == 'Y': signed_area += area
        support_val = signed_area / total_area if total_area else 0
    support = _pct_str(support_val)
    # SFR percentage: visible H11 if cached; otherwise basic count proxy.
    sfr = _pct_str(ws['H11'].value if ws['H11'].value not in [None,''] else 0)
    bl_values=[]
    for r in range(17,101):
        v = ws.cell(r,5).value
        try: v=float(v)
        except Exception: continue
        if v>0: bl_values.append(v)
    try: min_val=float(min_bl)
    except Exception: min_val=0
    qual = sum(1 for v in bl_values if v >= min_val)
    developed = len(bl_values)
    bl_pct = int(round((qual/developed)*100)) if developed else 0
    return {
        **basic,
        'Application_Number': app_no,
        'Support_Percentage': support,
        'SFR_Percentage': sfr,
        'Qualifying_Property_Count': qual,
        'Qualifying_Property_Count_Words': _num_words(qual),
        'Developed_Property_Count': developed,
        'Developed_Property_Count_Words': _num_words(developed),
        'Building_Line_Percentage': str(bl_pct),
        'Minimum_Building_Line_Feet': min_bl,
        'Minimum_Building_Line_Feet_Words': _num_words(min_bl),
        'Approval_Application_Line': f"The application is for the {basic.get('Block_Description','')}, {basic.get('Side_Of_Street','')}, {basic.get('Cross_Streets','')}.",
    }

def _clone_run_properties_from_paragraph(p):
    # Capture the formatting from the first meaningful run in the template paragraph.
    # This keeps template font, size, bold, underline, and color when we replace placeholder text.
    for r in p.runs:
        if r.text.strip():
            try:
                return deepcopy(r._r.get_or_add_rPr())
            except Exception:
                return None
    if p.runs:
        try:
            return deepcopy(p.runs[0]._r.get_or_add_rPr())
        except Exception:
            return None
    return None

def _replace_in_paragraph_preserve(p, replacements):
    text = p.text
    new = text
    for old,val in replacements.items():
        if old in new:
            new = new.replace(old, str(val))
    if new != text:
        template_rPr = _clone_run_properties_from_paragraph(p)
        p._p.clear_content()
        run = p.add_run(new)
        if template_rPr is not None:
            # Preserve font, size, bold, italic, underline, etc., but clear placeholder color
            # so generated text appears as normal black document text instead of red placeholder text.
            for child in list(template_rPr):
                if child.tag.endswith('}color'):
                    template_rPr.remove(child)
            rPr = run._r.get_or_add_rPr()
            run._r.remove(rPr)
            run._r.insert(0, deepcopy(template_rPr))

def _replace_doc_preserve(doc, replacements):
    for p in doc.paragraphs:
        _replace_in_paragraph_preserve(p, replacements)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_in_paragraph_preserve(p, replacements)

def generate_director_approval(template_path, wb_path, out_path):
    ci = read_case_info(wb_path); s2 = read_step2_approval(wb_path)
    doc = Document(template_path)
    minft = s2['Minimum_Building_Line_Feet']
    qual = s2['Qualifying_Property_Count']; dev = s2['Developed_Property_Count']
    application_line = s2['Approval_Application_Line']
    narrative = ci.get('Preservation_Finding') or f"Establishing a {minft}-foot minimum building line will preserve the building line character of the block."
    age_line = ' '.join([ci.get('Neighborhood_Age_Statement',''), ci.get('Structure_Age_Statement',''), narrative]).strip()
    replacements = {
        'Special Minimum Building Line Block No. XX': f"Special Minimum Building Line Block No. {s2['Application_Number']}",
        'The application is for the X block of X Street, X side, between X and X\nStreet.': application_line,
        'The application is for the X block of X Street, X side, between X and X Street.': application_line,
        'X% of the proposed application area is developed with not more than two SF \nresidential units per property.': f"{s2['SFR_Percentage']}% of the proposed application area is developed with not more than two SF residential units per property.",
        'X% of the proposed application area is developed with not more than two SF residential units per property.': f"{s2['SFR_Percentage']}% of the proposed application area is developed with not more than two SF residential units per property.",
        'A petition signed by owners of X% of the SMBLB.': f"A petition signed by owners of {s2['Support_Percentage']}% of the SMBLB.",
        'A minimum building line of X feet exists on X (X) properties in the blockface.': f"A minimum building line of {minft} feet exists on {s2['Qualifying_Property_Count_Words']} ({qual}) properties in the blockface.",
        'Establishing a X-foot minimum building line will preserve the building line \ncharacter of the block.': narrative,
        'Establishing a X-foot minimum building line will preserve the building line character of the block.': narrative,
        'The subdivision was platted in X. Most homes originated from the Xs.  The establishment of a X square foot minimum lot size will preserve the lot size character of the area.': age_line,
        'X (X) out of X (X) developed properties (representing X% of the application area) have a building line \nof at least X feet.': f"{s2['Qualifying_Property_Count_Words'].capitalize()} ({qual}) out of {s2['Developed_Property_Count_Words']} ({dev}) developed properties (representing {s2['Building_Line_Percentage']}% of the application area) have a building line of at least {minft} feet.",
        'X (X) out of X (X) developed properties (representing X% of the application area) have a building line of at least X feet.': f"{s2['Qualifying_Property_Count_Words'].capitalize()} ({qual}) out of {s2['Developed_Property_Count_Words']} ({dev}) developed properties (representing {s2['Building_Line_Percentage']}% of the application area) have a building line of at least {minft} feet.",
        'Vonn Tran': ci.get('Director_Name','Vonn Tran'),
    }
    _replace_doc_preserve(doc, replacements)
    remove_trailing_blank_paragraphs(doc)
    doc.save(out_path)
    unlock_docx(out_path)  # fully editable output
    return out_path

def generate_pc_approval(template_path, wb_path, out_path):
    ci = read_case_info(wb_path); s2 = read_step2_approval(wb_path)
    doc = Document(template_path)
    minft=s2['Minimum_Building_Line_Feet']; qual=s2['Qualifying_Property_Count']; dev=s2['Developed_Property_Count']
    application_line=s2['Approval_Application_Line']
    narrative = ci.get('Preservation_Finding') or f"The establishment of a {minft} foot minimum building line will preserve the building line character of the area."
    age_line=' '.join([ci.get('Neighborhood_Age_Statement',''), ci.get('Structure_Age_Statement',''), narrative]).strip()
    signer=ci.get('PC_Signer','Martha Stein, Chair')
    # The PC template contains two signature alternatives. Replace Chair line with selected signer and keep the alternate visible as in template.
    replacements={
        'Special Building Line Requirement Area No. 106': f"Special Building Line Requirement Area No. {s2['Application_Number']}",
        'The application is for the 300 block of East 24th Street, north and south sides.': application_line,
        '88% of the proposed application area is developed with single-family \nresidential units per property.': f"{s2['SFR_Percentage']}% of the proposed application area is developed with single-family residential units per property.",
        '88% of the proposed application area is developed with single-family residential units per property.': f"{s2['SFR_Percentage']}% of the proposed application area is developed with single-family residential units per property.",
        'Petition signed by owners of 53% of the SMBLB.': f"Petition signed by owners of {s2['Support_Percentage']}% of the SMBLB.",
        'A minimum building line of 15 feet exists on ten (10) properties in the \nblockface.': f"A minimum building line of {minft} feet exists on {s2['Qualifying_Property_Count_Words']} ({qual}) properties in the blockface.",
        'A minimum building line of 15 feet exists on ten (10) properties in the blockface.': f"A minimum building line of {minft} feet exists on {s2['Qualifying_Property_Count_Words']} ({qual}) properties in the blockface.",
        'The subdivision was platted in xxxx. The houses mostly originate from the \n19xxs. The establishment of a 15 foot minimum building line will preserve the \nbuilding line character of the area.': age_line,
        'The subdivision was platted in xxxx.  The houses mostly originate from the 19xxs.  The establishment of a 15 foot minimum building line will preserve the building line character of the area.': age_line,
        'The subdivision was platted in xxxx. The houses mostly originate from the 19xxs. The establishment of a 15 foot minimum building line will preserve the building line character of the area.': age_line,
        'Ten (10) out of thirteen (13) developed properties (representing 77% of the application area) have a \nbuilding line of at least 15 feet.': f"{s2['Qualifying_Property_Count_Words'].capitalize()} ({qual}) out of {s2['Developed_Property_Count_Words']} ({dev}) developed properties (representing {s2['Building_Line_Percentage']}% of the application area) have a building line of at least {minft} feet.",
        'Ten (10) out of thirteen (13) developed properties (representing 77% of the application area) have a building line of at least 15 feet.': f"{s2['Qualifying_Property_Count_Words'].capitalize()} ({qual}) out of {s2['Developed_Property_Count_Words']} ({dev}) developed properties (representing {s2['Building_Line_Percentage']}% of the application area) have a building line of at least {minft} feet.",
        'Martha Stein, Chair': signer,
        'Martha L. Stein, Chair': signer,
    }
    _replace_doc_preserve(doc, replacements)
    remove_trailing_blank_paragraphs(doc)
    doc.save(out_path)
    unlock_docx(out_path)  # fully editable output
    return out_path


# ---- MBL routing form generation ----
def _iter_sdt(doc):
    return list(doc.element.xpath('.//w:sdt'))

def _sdt_text(sdt):
    try:
        return ''.join(t.text or '' for t in sdt.xpath('.//*[local-name()="t"]'))
    except Exception:
        return ''

def _set_sdt_text(sdt, value):
    # Preserve the content control and its run formatting; replace only visible text.
    texts = list(sdt.xpath('.//*[local-name()="t"]'))
    if not texts:
        return
    texts[0].text = str(value)
    for t in texts[1:]:
        t.text = ''
    # Clear red placeholder color if present on generated controls.
    for color in sdt.xpath('.//*[local-name()="color"]'):
        try:
            color.getparent().remove(color)
        except Exception:
            pass

def _set_checkbox_sdt(sdt, checked):
    symbol = '☒' if checked else '☐'
    # Set the checkbox control state.
    for chk in sdt.xpath('.//*[local-name()="checked"]'):
        try:
            chk.set('{http://schemas.microsoft.com/office/word/2010/wordml}val', '1' if checked else '0')
        except Exception:
            try:
                chk.set('w14:val', '1' if checked else '0')
            except Exception:
                pass
    # Set visible symbol.
    texts=list(sdt.xpath('.//*[local-name()="t"]'))
    if texts:
        texts[0].text=symbol
        for t in texts[1:]: t.text=''

def _replace_plain_runs(doc, replacements):
    # Low-risk text-node replacement. Keeps paragraph/runs and content controls intact.
    for t in doc.element.xpath('.//*[local-name()="t"]'):
        if t.text:
            new=t.text
            for old,val in replacements.items():
                if old in new:
                    new=new.replace(old, str(val))
            t.text=new



def _clear_text_colors_in_doc(doc):
    """Clear direct run/paragraph text colors in generated forms so red template placeholders
    become normal black text while preserving borders, layout, bolding, and spacing."""
    for color in doc.element.xpath('.//*[local-name()="color"]'):
        try:
            color.getparent().remove(color)
        except Exception:
            pass

def _set_paragraph_left_indent_for_sdt(sdt, left_twips=2160):
    """Force wrapped lines in a content-control paragraph to align under the first line."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    p = sdt
    while p is not None and p.tag.split('}')[-1] != 'p':
        p = p.getparent()
    if p is None:
        return
    pPr = p.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        p.insert(0, pPr)
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:left'), str(left_twips))
    # Remove firstLine so wrapped lines align with the first line.
    for attr in (qn('w:firstLine'), qn('w:hanging')):
        if attr in ind.attrib:
            del ind.attrib[attr]

def generate_routing_form(template_path, wb_path, out_path):
    ci = read_case_info(wb_path)
    s2 = read_step2_approval(wb_path)
    doc = Document(template_path)
    app_no = s2.get('Application_Number','')
    min_bl = s2.get('Minimum_Building_Line_Feet','')
    support = s2.get('Support_Percentage','')
    sfr = s2.get('SFR_Percentage','')
    qual = s2.get('Qualifying_Property_Count','')
    dev = s2.get('Developed_Property_Count','')
    # Topic/area strings.
    topic = f"Special Minimum Building Line Block Application No. {app_no}"
    area = f"The {s2.get('Block_Description','')}, {s2.get('Side_Of_Street','')}, {s2.get('Cross_Streets','')}"
    author = f"{ci.get('Planner_Name','')}, {ci.get('Planner_Phone','')}".strip().strip(',')
    council_deadline = ci.get('Council_Deadline','')
    approval_path = ci.get('Approval_Path','Director Approval')
    is_pc = approval_path.lower().startswith('planning')
    # Replace the simple red placeholders and labels while preserving the template structure.
    _replace_plain_runs(doc, {
        'Special Minimum Lot Size/ Block Application No. X': topic,
        'Special Minimum Lot Size Block Application No. X': topic,
        'The X block of X Street, X and X sides, between X Street and X Street': area,
        'Planner Name, Phone Number': author,
        'June 11, 2026': council_deadline,
        'Lot Size/Building Line Established': 'Building Line Established',
        'Lot Size/': '',
        'Total Lots /': 'Total Developed Properties /',
        'that meet Lot Size': 'that meet Building Line',
    })
    sdts = _iter_sdt(doc)
    # Fill top content controls directly; some placeholder text is stored inside Word content controls.
    for idx,val in {0: topic, 2: area, 3: author, 4: council_deadline}.items():
        if len(sdts)>idx:
            _set_sdt_text(sdts[idx], val)
    if len(sdts) > 2:
        _set_paragraph_left_indent_for_sdt(sdts[2], left_twips=2160)
    # Fix the MBL count label without globally changing every occurrence of 'Lot'.
    for t in doc.element.xpath('.//*[local-name()="t"]'):
        if t.text == 'Lot ':
            t.text = 'Properties '
    # Set known text content controls by order from the uploaded routing template.
    # 5-6 are Building Line / Lot Size checkboxes.
    if len(sdts) > 6:
        _set_checkbox_sdt(sdts[5], True)   # Building Line
        _set_checkbox_sdt(sdts[6], False)  # Lot Size
    # 7 app no, 8 established building line, 15 SFR, 17 support, 19 counts.
    mapping = {
        7: str(app_no),
        8: f': {min_bl} feet',
        15: str(sfr),
        17: f' {support}',
        19: f'{dev} / {qual}',
    }
    for idx,val in mapping.items():
        if len(sdts)>idx:
            _set_sdt_text(sdts[idx], val)
    # RCA Applies to checkboxes.
    # 20 SMLS Area, 21 SMLS Block, 22 SMBL Block, 23 SMLS Renewal, 24 SMBL Renewal
    for idx,checked in [(20,False),(21,False),(22,True),(23,False),(24,False)]:
        if len(sdts)>idx: _set_checkbox_sdt(sdts[idx], checked)
    # Attachments Provided.
    # 25 PC Approval, 26 Director Approval, 27 Map, 28 Staff Analysis, 29 Plat, 30 Tax Map, 31 Tax Records, 32 Original Ordinance
    attachment_states = {
        25: is_pc,
        26: not is_pc,
        27: True,
        28: True,
        29: True,
        30: True,
        31: True,
        32: False,
    }
    for idx,checked in attachment_states.items():
        if len(sdts)>idx: _set_checkbox_sdt(sdts[idx], checked)
    _clear_text_colors_in_doc(doc)
    remove_trailing_blank_paragraphs(doc)
    prepare_routing_form(doc)
    doc.save(out_path)
    unlock_docx(out_path)  # fully editable output
    return out_path


# ---- MBL RCA Packet Builder ----
def _find_soffice() -> str | None:
    for exe in ["soffice", "libreoffice"]:
        found = shutil.which(exe)
        if found:
            return found
    for p in [
        r"C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        r"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    ]:
        if Path(p).exists():
            return p
    return None

def convert_office_to_pdf(input_path, output_dir):
    input_path=Path(input_path); output_dir=Path(output_dir); output_dir.mkdir(parents=True, exist_ok=True)
    # LibreOffice when installed, otherwise Microsoft Word/Excel (City PCs).
    from core.office_convert import office_to_pdf
    return office_to_pdf(input_path, output_dir)

def image_to_pdf(input_path, output_path):
    from PIL import Image
    input_path=Path(input_path); output_path=Path(output_path); output_path.parent.mkdir(parents=True, exist_ok=True)
    img=Image.open(input_path)
    if img.mode in ("RGBA","LA"):
        bg=Image.new("RGB", img.size, "white"); bg.paste(img, mask=img.split()[-1]); img=bg
    else:
        img=img.convert("RGB")
    img.save(output_path, "PDF", resolution=150)
    return output_path

def ensure_pdf(input_path, work_dir):
    input_path=Path(input_path); work_dir=Path(work_dir); suf=input_path.suffix.lower()
    if suf=='.pdf': return input_path
    if suf in ('.docx','.doc','.xlsx','.xls'): return convert_office_to_pdf(input_path, work_dir)
    if suf in ('.png','.jpg','.jpeg','.tif','.tiff'): return image_to_pdf(input_path, work_dir/f"{input_path.stem}.pdf")
    raise ValueError(f"Unsupported packet attachment type: {input_path}")

def _get_pdf_writer_class():
    try:
        from pypdf import PdfWriter
        return PdfWriter
    except ModuleNotFoundError:
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pypdf>=4.0.0'])
            from pypdf import PdfWriter
            return PdfWriter
        except Exception as exc:
            raise ModuleNotFoundError("The RCA Packet PDF builder needs the 'pypdf' package. Run install_requirements.bat, then try again.") from exc

def merge_pdfs(pdf_paths, output_pdf):
    PdfWriter=_get_pdf_writer_class(); writer=PdfWriter()
    for p in pdf_paths:
        p=Path(p)
        if not p.exists(): raise FileNotFoundError(p)
        writer.append(str(p))
    output_pdf=Path(output_pdf); output_pdf.parent.mkdir(parents=True, exist_ok=True)
    with open(output_pdf,'wb') as f: writer.write(f)
    return output_pdf

def _blankish_staff_value(v):
    """Treat formula/helper output as blank for Step 2 print-range detection."""
    if v is None: return True
    if isinstance(v, (int, float)):
        return float(v) == 0
    s = str(v).strip()
    return s in ('', '0', '0.0', '#NAME?', '#N/A', '###', 'None')

def _detect_mbl_step2_staff_range(excel_path):
    """Detect the official MBL Step 2 staff-analysis print range.

    The MBL calculation sheet often contains helper formulas far below the actual property
    table. Those rows may display 0s, ranks, or formula-error placeholders and should not
    be printed. The packet should include the top calculation section plus the property rows
    where the left-side official table has real property information.
    """
    wb_vals=openpyxl.load_workbook(excel_path, data_only=True)
    sheet='Step 2 - Main Calculation Sheet'
    if sheet not in wb_vals.sheetnames: raise ValueError(f"Could not find required sheet: {sheet}")
    ws=wb_vals[sheet]

    # Always keep the header/calculation area through the property table heading.
    last_row=16

    # Find real property rows using the official left-side table, not the helper/rank columns.
    # A row is real if it has an actual property description/land use, or a nonzero lot/building-line value.
    for r in range(17, ws.max_row+1):
        prop = ws.cell(r,1).value   # Property Description
        land = ws.cell(r,2).value   # Land Use
        lot  = ws.cell(r,4).value   # Lot size
        bl   = ws.cell(r,5).value   # Building Line
        has_prop_text = not _blankish_staff_value(prop) or not _blankish_staff_value(land)
        has_real_measure = not _blankish_staff_value(lot) or not _blankish_staff_value(bl)
        if has_prop_text or has_real_measure:
            last_row=r

    # Use only the official visible staff-analysis columns. Never include helper columns to the right.
    return f"A1:H{last_row}"

def export_mbl_step2_staff_analysis_pdf(excel_path, output_pdf):
    excel_path=Path(excel_path); output_pdf=Path(output_pdf)
    tmpdir=Path(tempfile.mkdtemp(prefix='mbl_step2_'))
    tmp_xlsx=tmpdir/'MBL_Staff_Analysis_Step2.xlsx'
    wb=openpyxl.load_workbook(excel_path, data_only=False)
    wb_vals=openpyxl.load_workbook(excel_path, data_only=True)
    sheet='Step 2 - Main Calculation Sheet'
    if sheet not in wb.sheetnames or sheet not in wb_vals.sheetnames:
        raise ValueError(f"Could not find required sheet: {sheet}")
    ws=wb[sheet]; ws_vals=wb_vals[sheet]
    print_area=_detect_mbl_step2_staff_range(excel_path)
    from openpyxl.utils.cell import range_boundaries
    min_col,min_row,max_col,max_row=range_boundaries(print_area)
    for row in range(min_row, max_row+1):
        for col in range(min_col, max_col+1):
            cell=ws.cell(row,col); val_cell=ws_vals.cell(row,col)
            if isinstance(cell.value,str) and cell.value.startswith('='):
                cell.value=val_cell.value
    # Clear outside official A:H print area so helper columns/blank cells do not appear.
    for row in range(1, ws.max_row+1):
        for col in range(1, ws.max_column+1):
            if row<min_row or row>max_row or col<min_col or col>max_col:
                
                try:
                    ws.cell(row,col).value=None
                except AttributeError:
                    pass
    for s in list(wb.sheetnames):
        if s!=sheet: del wb[s]
    ws=wb[sheet]
    apply_staff_analysis_page_setup(ws, print_area, min_row, max_row, min_col, max_col, header_row=16)
    wb.save(tmp_xlsx)
    pdf=convert_office_to_pdf(tmp_xlsx,tmpdir)
    output_pdf.parent.mkdir(parents=True, exist_ok=True); shutil.copyfile(pdf, output_pdf)
    return output_pdf

def generate_mbl_packet_pdf(wb_path, templates, map_path, subdivision_plat_path, tax_map_path, hcad_records_path, output_pdf):
    """Generate the MBL RCA packet PDF. Planner supplies only plat, tax map, and one/multiple HCAD PDFs; the wizard generates the rest."""
    wb_path=Path(wb_path); output_pdf=Path(output_pdf); map_path=Path(map_path)
    ci=read_case_info(wb_path)
    approval_path=ci.get('Approval_Path','Director Approval')
    is_pc=approval_path.lower().startswith('planning')
    from core.packet_work import packet_workspace
    with packet_workspace() as work_dir:
        pdfs=[]
        routing_docx=work_dir/'01_MBL_Routing_Form.docx'
        generate_routing_form(templates['routing'], wb_path, routing_docx)
        pdfs.append(routing_pdf_one_page(ensure_pdf(routing_docx, work_dir), work_dir))
        if is_pc:
            pc_docx=work_dir/'02_MBL_PC_Approval.docx'
            generate_pc_approval(templates['pc'], wb_path, pc_docx)
            pdfs.append(ensure_pdf(pc_docx, work_dir))
        else:
            director_docx=work_dir/'02_MBL_Director_Approval.docx'
            generate_director_approval(templates['director'], wb_path, director_docx)
            pdfs.append(ensure_pdf(director_docx, work_dir))
        pdfs.append(ensure_pdf(map_path, work_dir))
        staff_pdf=work_dir/'04_MBL_Staff_Analysis_Step2.pdf'
        export_mbl_step2_staff_analysis_pdf(wb_path, staff_pdf)
        pdfs.append(staff_pdf)
        if str(subdivision_plat_path or "").strip():
            pdfs.append(ensure_pdf(subdivision_plat_path, work_dir))
        pdfs.append(ensure_pdf(tax_map_path, work_dir))
        if isinstance(hcad_records_path,(list,tuple)):
            hcad_paths=[Path(p) for p in hcad_records_path if str(p).strip()]
        else:
            hcad_paths=[Path(p) for p in str(hcad_records_path or '').split('||') if p.strip()]
        # Each HCAD record is condensed onto one page so a record that spills onto a
        # second page does not split across the packet.
        pdfs.extend(condense_hcad_pdfs([ensure_pdf(h, work_dir) for h in hcad_paths], work_dir))
        merge_pdfs(pdfs, output_pdf)
        return len(pdfs)

# ---- MBL Draft Ordinance Generator ----
def _mbl_read_step1_properties(wb_path):
    import re
    wb = openpyxl.load_workbook(wb_path, data_only=True)
    ws = wb['Step 1 - Input Properties Info']
    hm = header_map(ws)
    def gv(row, name):
        c = hm.get(name)
        return clean(ws.cell(row, c).value) if c else ''
    props=[]
    seen=set()
    for r in range(2, ws.max_row+1):
        addr = ' '.join([gv(r,'Site_addr_1'), gv(r,'Site_addr_2'), gv(r,'Site_addr_3')]).strip()
        owner = gv(r,'OWNER_MAILTO')
        hcad = gv(r,'TAX_ID') or gv(r,'ALT_TAX_ID')
        legal_parts=[gv(r,'Legal_Dscr_1'), gv(r,'Legal_Dscr_2'), gv(r,'Legal_Dscr_3'), gv(r,'Legal_Dscr_4')]
        legal=' '.join([p for p in legal_parts if p and p not in ['0','None']]).strip()
        propdesc = gv(r,'Property Description') or legal
        if not addr and not legal and not hcad: continue
        key=(addr,hcad,legal)
        if key in seen: continue
        seen.add(key)
        text = propdesc or legal
        lottract=''
        block=''
        subdivision=''
        # Best-effort parse from common HCAD legal-description phrasing.
        m=re.search(r'((?:LT|LTS|TR|TRS)\.?\s*[^,;]*?)(?:\s+BLK\b|$)', text, re.I)
        if m: lottract=m.group(1).strip()
        m=re.search(r'\bBLK\.?\s*([A-Za-z0-9\-]+)', text, re.I)
        if m: block=m.group(1).strip()
        # Use the last legal descriptor that is not obviously lot/block as subdivision fallback.
        for part in reversed(legal_parts):
            if part and not re.search(r'\b(LT|LTS|TR|TRS|BLK)\b', part, re.I):
                subdivision=part.strip(); break
        if not subdivision:
            # Remove leading lot/block phrases to get a fallback subdivision.
            subdivision=re.sub(r'^(?:LT|LTS|TR|TRS).*?\bBLK\.?\s*[A-Za-z0-9\-]+\s*', '', text, flags=re.I).strip()
        props.append({'address':addr,'hcad':hcad,'legal':legal or propdesc,'lottract':lottract,'block':block,'subdivision':subdivision})
    return props

def _mbl_build_ordinance_context(wb_path):
    ci=read_case_info(wb_path)
    s2=read_step2_approval(wb_path)
    area_lower=f"{s2.get('Block_Description','')}, {s2.get('Side_Of_Street','')}, {s2.get('Cross_Streets','')}".replace(' ,',',').strip().strip(',')
    context={
        'Ordinance_Year': ci.get('Ordinance_Year') or '2026',
        'Ordinance_Area_Upper': area_lower.upper(),
        'Ordinance_Area_Lower': area_lower,
        'Support_Percentage': s2.get('Support_Percentage',''),
        # Reuse the existing cleaned ordinance placeholders so the MBL template can be created from the MLS template.
        'Qualifying_Lot_Count_Words': s2.get('Qualifying_Property_Count_Words',''),
        'Qualifying_Lot_Count': s2.get('Qualifying_Property_Count',''),
        'Total_Lot_Count_Words': s2.get('Developed_Property_Count_Words',''),
        'Total_Lot_Count': s2.get('Developed_Property_Count',''),
        'Minimum_Lot_Size_Words': s2.get('Minimum_Building_Line_Feet_Words',''),
        'Minimum_Lot_Size_Number': s2.get('Minimum_Building_Line_Feet',''),
        'Director_Name': ci.get('Director_Name','Vonn Tran'),
        'Legal_File_No': ci.get('Legal_File_No','__________'),
    }
    return context

def _replace_placeholders_in_doc(doc, context):
    # For ordinance body text, replace placeholder tokens inside existing runs so
    # bolding/spacing from the City template is preserved. Do not rewrite whole paragraphs.
    repl={f'{{{{{k}}}}}':v for k,v in context.items()}
    _replace_plain_runs(doc, repl)

def _insert_paragraph_after(paragraph, text=None, style=None):
    from docx.text.paragraph import Paragraph
    from docx.oxml import OxmlElement
    new_p = OxmlElement('w:p')
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        run=new_para.add_run(text)
    if style:
        new_para.style=style
    return new_para

def _insert_table_after(paragraph, rows, cols):
    from docx.table import Table
    tbl = OxmlElement('w:tbl')
    paragraph._p.addnext(tbl)
    table = Table(tbl, paragraph._parent)
    # python-docx cannot easily build an empty table this way, so use parent doc style fallback by inserting through a temp doc is hard.
    return table

def _add_exhibit_c_table(doc, props):
    # Add property list table immediately after the Exhibit C heading if possible, otherwise append at end.
    target=None
    for p in doc.paragraphs:
        if 'PROPERTY LIST' in p.text:
            target=p
            break
    table=doc.add_table(rows=1, cols=5)
    table.style='Table Grid'
    hdr=table.rows[0].cells
    for i,h in enumerate(['ADDRESS','HCAD','LOT/TRACT','BLOCK','SUBDIVISION']):
        hdr[i].text=h
    for pr in props:
        row=table.add_row().cells
        row[0].text=pr.get('address','')
        row[1].text=pr.get('hcad','')
        row[2].text=pr.get('lottract','')
        row[3].text=pr.get('block','')
        row[4].text=pr.get('subdivision','')
    # Move table after the target paragraph by relocating its XML.
    if target is not None:
        tbl_el=table._tbl
        tbl_el.getparent().remove(tbl_el)
        target._p.addnext(tbl_el)
    return table

def _fill_exhibit_a(doc, props):
    legal_lines=[]
    seen=set()
    for pr in props:
        legal=pr.get('legal','').strip()
        if legal and legal not in seen:
            seen.add(legal); legal_lines.append(legal)
    text='\n'.join(legal_lines)
    if not text: text=''
    replaced=False
    for p in doc.paragraphs:
        if 'Click or tap here to enter text.' in p.text:
            p._p.clear_content()
            run=p.add_run(text)
            try: run.font.color.rgb=RGBColor(0,0,0)
            except Exception: pass
            replaced=True
            break
    if not replaced:
        for p in doc.paragraphs:
            if 'PROPERTY DESCRIPTIONS' in p.text:
                _insert_paragraph_after(p, text)
                break

def _insert_exhibit_b_map(doc, map_path):
    target=None
    for p in doc.paragraphs:
        if 'BOUNDARY MAP' in p.text:
            target=p
            break
    if target is None: return
    with tempfile.TemporaryDirectory() as td:
        img=map_image_for_docx(map_path, td)
        newp=_insert_paragraph_after(target, '')
        run=newp.add_run()
        run.add_picture(img, width=Inches(6.0))

def generate_draft_ordinance(wb_path, template_path, map_path, out_path):
    """Generate a planner-draft MBL ordinance using the MBL ordinance template.
    Legal blanks remain intact. The wizard fills only planning facts and exhibits.
    """
    context=_mbl_build_ordinance_context(wb_path)
    props=_mbl_read_step1_properties(wb_path)
    doc=Document(template_path)
    _replace_placeholders_in_doc(doc, context)
    _fill_exhibit_a(doc, props)
    _insert_exhibit_b_map(doc, map_path)
    _add_exhibit_c_table(doc, props)
    _clear_text_colors_in_doc(doc)
    remove_trailing_blank_paragraphs(doc)
    doc.save(out_path)
    unlock_docx(out_path)  # fully editable output
    return len(props)


# ---------------- HCAD records export (ported from MLS wizard) ----------------
import re as _re_hcad


def collect_hcad_records_from_step1(wb_path):
    """One HCAD property record per normalized account, including separated lots."""
    from core.hcad_identity import collect_hcad_records
    return collect_hcad_records(wb_path)



def export_hcad_address_list(wb_path, output_folder):
    """Write Step 1 HCAD search records to Output/HCAD_Records/MBL_<app#>_HCAD_Address_List.txt.

    The file is tab-delimited (ADDRESS<TAB>TAX_ID) so the downloader can search by
    address first and fall back to a cleaned account number.
    """
    ci = read_case_info(wb_path)
    app_no = _re_hcad.sub(r'[^A-Za-z0-9_-]+', '', ci.get('Application_Number', '')) or 'Case'
    hcad_folder = Path(output_folder) / 'HCAD_Records'
    hcad_folder.mkdir(parents=True, exist_ok=True)
    address_path = hcad_folder / f'MBL_{app_no}_HCAD_Address_List.txt'
    records = collect_hcad_records_from_step1(wb_path)
    if not records:
        raise ValueError('No site addresses or TAX_ID values were found on Step 1 for HCAD download.')
    lines = []
    for rec in records:
        address = rec.get('Address', '')
        tax_id = rec.get('TAX_ID', '')
        lines.append(f"{address}\t{tax_id}" if tax_id else address)
    address_path.write_text('\n'.join(lines) + '\n', encoding='utf-8')
    return address_path, len(records)
