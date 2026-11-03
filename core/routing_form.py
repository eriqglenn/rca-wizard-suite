"""Keep the City's routing checklist together on one editable Letter page."""
from __future__ import annotations

from copy import deepcopy
from pathlib import Path

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def _text(element):
    return ''.join(t.text or '' for t in element.iter(qn('w:t'))).strip()


def _property(parent, tag, **attrs):
    element = parent.find(qn(tag))
    if element is None:
        element = OxmlElement(tag)
        parent.append(element)
    for name, value in attrs.items():
        element.set(qn('w:' + name), str(value))
    return element


def _line_paragraphs(paragraph):
    """Split checklist line breaks while retaining runs and checkbox controls."""
    ppr = paragraph.find(qn('w:pPr'))

    def blank():
        p = OxmlElement('w:p')
        if ppr is not None:
            props = deepcopy(ppr)
            for tag in ('w:sectPr', 'w:keepNext', 'w:pageBreakBefore'):
                for node in props.findall(qn(tag)):
                    props.remove(node)
            p.append(props)
        props = p.find(qn('w:pPr'))
        if props is None:
            props = OxmlElement('w:pPr')
            p.append(props)
        _property(props, 'w:spacing', before=0, after=0, line=240, lineRule='auto')
        return p

    lines, current = [], blank()
    for child in paragraph:
        if child.tag == qn('w:pPr'):
            continue
        if child.tag != qn('w:r') or child.find(qn('w:br')) is None:
            current.append(deepcopy(child))
            continue
        rpr = child.find(qn('w:rPr'))
        run = OxmlElement('w:r')
        if rpr is not None:
            run.append(deepcopy(rpr))
        for item in child:
            if item.tag == qn('w:rPr'):
                continue
            if item.tag == qn('w:br'):
                if len(run) > (1 if rpr is not None else 0):
                    current.append(run)
                if _text(current):
                    lines.append(current)
                current, run = blank(), OxmlElement('w:r')
                if rpr is not None:
                    run.append(deepcopy(rpr))
            else:
                run.append(deepcopy(item))
        if len(run) > (1 if rpr is not None else 0):
            current.append(run)
    if _text(current):
        lines.append(current)
    return lines


def _checklist_table(doc, heading, stop_heading=None):
    body = doc.element.body
    children = list(body)
    start = next((i for i, p in enumerate(children)
                  if p.tag == qn('w:p') and _text(p).startswith(heading)), None)
    if start is None:
        return
    end = next((i for i in range(start + 1, len(children))
                if children[i].tag == qn('w:sectPr') or
                (stop_heading and _text(children[i]).startswith(stop_heading))), len(children))
    originals = children[start + 1:end]
    # A custom form with an existing table already has stable cell placement.
    if any(p.tag != qn('w:p') for p in originals):
        return
    lines = [line for p in originals for line in _line_paragraphs(p)]
    if not lines:
        return

    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    props = table._tbl.tblPr
    _property(props, 'w:tblW', w=10080, type='dxa')
    _property(props, 'w:tblInd', w=0, type='dxa')
    borders = _property(props, 'w:tblBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        _property(borders, 'w:' + side, val='nil')
    trpr = table.rows[0]._tr.get_or_add_trPr()
    _property(trpr, 'w:cantSplit')
    for index, (cell, items) in enumerate(zip(table.rows[0].cells, (lines[:3], lines[3:]))):
        for child in list(cell._tc):
            if child.tag != qn('w:tcPr'):
                cell._tc.remove(child)
        tcpr = cell._tc.get_or_add_tcPr()
        _property(tcpr, 'w:tcW', w=5040, type='dxa')
        margins = _property(tcpr, 'w:tcMar')
        for side, value in [('top', 0), ('bottom', 0), ('left', 360 if index else 0), ('right', 0 if index else 360)]:
            _property(margins, 'w:' + side, w=value, type='dxa')
        for p in items:
            cell._tc.append(p)
        if not items:
            cell._tc.append(OxmlElement('w:p'))
    children[start].addnext(table._tbl)
    for p in originals:
        body.remove(p)
    ppr = children[start].get_or_add_pPr()
    _property(ppr, 'w:keepNext')
    _property(ppr, 'w:spacing', before=100 if heading == 'Attachments Provided:' else 0,
              after=40, line=240, lineRule='auto')


def prepare_routing_form(doc):
    """Replace fragile newspaper columns with the same two-column checklist.

    Text, fonts, signature lines, and editable checkbox states are preserved.
    Only the checklist layout and excess section spacing change.
    """
    _checklist_table(doc, 'RCA Applies to:', 'Attachments Provided:')
    _checklist_table(doc, 'Attachments Provided:')
    body = doc.element.body
    for ppr in list(body.iter(qn('w:pPr'))):
        for tag in ('w:sectPr', 'w:pageBreakBefore'):
            for node in list(ppr.findall(qn(tag))):
                ppr.remove(node)
    for node in list(body.iter(qn('w:lastRenderedPageBreak'))):
        node.getparent().remove(node)
    for node in list(body.iter(qn('w:br'))):
        if node.get(qn('w:type')) in ('page', 'column'):
            node.getparent().remove(node)
    section = body.find(qn('w:sectPr'))
    if section is None:
        section = OxmlElement('w:sectPr')
        body.append(section)
    _property(section, 'w:pgSz', w=12240, h=15840)
    _property(section, 'w:pgMar', top=1440, bottom=720, left=1080, right=1080,
              header=720, footer=360, gutter=0)
    _property(section, 'w:cols', num=1, space=720)
    # Word needs an ending paragraph after a table; reserve only one point.
    while section.getprevious() is not None:
        previous = section.getprevious()
        if previous.tag != qn('w:p') or _text(previous) or any(
                list(previous.iter(qn(tag))) for tag in
                ('w:drawing', 'w:pict', 'w:object', 'w:fldChar', 'w:instrText', 'w:sdt')):
            break
        body.remove(previous)
    tail = OxmlElement('w:p')
    ppr = OxmlElement('w:pPr')
    _property(ppr, 'w:spacing', before=0, after=0, line=20, lineRule='exact')
    tail.append(ppr)
    section.addprevious(tail)


def routing_pdf_one_page(pdf_path, work_dir):
    """Keep a converter's overflow on the routing sheet before packet assembly."""
    from .pdf_condense import condense_pdf_to_one_page
    pdf_path = Path(pdf_path)
    return condense_pdf_to_one_page(pdf_path, Path(work_dir) / ('one_page_' + pdf_path.name))
