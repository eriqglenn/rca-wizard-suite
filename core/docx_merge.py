"""Shared Word mail-merge helpers for the City templates (2026 wording-based set).

Everything here replaces case-specific placeholder text inside the existing
runs / content controls of an official City template and never rebuilds a
form.  Non-breaking spaces (the templates are full of them) are normalised
before matching.  Extracted from the MLS Renewal wizard so the MLS wizard can
share the same engine; the renewal wizard keeps its own verified copy.
"""
import datetime
import re
import tempfile
from copy import deepcopy
from pathlib import Path

import fitz
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docx.text.paragraph import Paragraph

_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'


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


_PHONE_RE = re.compile(r'\(?832\)?[ -]?393-?')


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
                if nt and set(nt) <= set('Xx-'):
                    _set_run_text_preserve_xml(nxt, '')
                elif nt:
                    break
            return True
    return False


_EMAIL_PLACEHOLDERS = ['Planner Name@houstontx.gov', 'PlannerName@houstontx.gov', 'Plannner@houstontx.gov', 'Planner.Name@houstontx.gov', 'Name@houstontx.gov']


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


