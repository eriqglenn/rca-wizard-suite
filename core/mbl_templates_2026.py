"""Fill the September 2026 City MBL forms without rebuilding their layout."""
from decimal import Decimal, InvalidOperation
from docx.oxml.ns import qn

PREFIX = '{{MBL2026_'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'


def _checkboxes(paragraph, states):
    """Support both Word checkbox controls and literal City checkbox glyphs."""
    nodes = [n for n in paragraph.iter(qn('w:t')) if any(c in (n.text or '') for c in '☐☒☑')]
    choices = iter(states)
    for node in nodes:
        chars = []
        for char in node.text:
            if char not in '☐☒☑':
                chars.append(char)
                continue
            state = next(choices, None)
            chars.append(char if state is None else ('☒' if state else '☐'))
            if state is not None:
                ancestor = node.getparent()
                while ancestor is not None and ancestor is not paragraph:
                    if ancestor.tag == qn('w:sdt'):
                        for checked in ancestor.iter('{'+W14+'}checked'):
                            checked.set('{'+W14+'}val', '1' if state else '0')
                        break
                    ancestor = ancestor.getparent()
        node.text = ''.join(chars)


def sfr_threshold(context):
    """Use counts, not the displayed rounded percentage; unknown is not No."""
    try:
        total = Decimal(str(context.get('Total_Lot_Count', '')))
        count = Decimal(str(context.get('SFR_Count', context.get('SFR_Lot_Count', ''))))
        if not total.is_finite() or not count.is_finite() or total <= 0 or not 0 <= count <= total:
            return None
        return count * 100 >= total * 60
    except (InvalidOperation, ValueError, TypeError):
        return None


def fill_mbl_template(doc, context, *, renewal, routing=False):
    """Return False for older/custom templates so their original filler remains."""
    nodes = list(doc.element.iter(qn('w:t')))
    if not any(PREFIX in (n.text or '') for n in nodes):
        return False
    def value(key, fallback=''):
        raw = context.get(key)
        return str(raw).strip() if raw is not None and str(raw).strip() else fallback
    number = value('Application_Renewal_Number' if renewal else 'Application_Number', '[APPLICATION NUMBER]')
    fields = {
        'NUMBER': number,
        'TOPIC': f"Special Minimum Building Line Block {'Renewal ' if renewal else ''}Application No. {number}",
        'AREA': value('Area_Description', '[APPLICATION AREA]'),
        'LINE': value('Minimum_Building_Line_Text', '[BUILDING LINE]'),
        'DIRECTOR': value('Director_Name', 'Vonn Tran'),
        'APPROVAL_DATE': value('Director_Approval_Date', 'Date'),
        'PLANNER': ', '.join(filter(None, [value('Planner_Name'), value('Planner_Phone')])),
        'DEADLINE': value('Council_Deadline', 'DATE'),
        'ORDINANCE': value('Original_Ordinance_No', '[PREVIOUS ORDINANCE]') if renewal else 'N/A',
        'DISTRICT': value('Council_District', '[DISTRICT]'),
    }
    for node in nodes:
        for key, replacement in fields.items():
            node.text = (node.text or '').replace('{{MBL2026_'+key+'}}', replacement)
    # Content-control runs are omitted by Document.paragraphs/runs. Include
    # those runs when clearing the City's red placeholder formatting.
    for color in doc.element.iter(qn('w:color')):
        shade = color.get(qn('w:val'), '').upper()
        if len(shade) == 6 and shade.endswith('0000') and shade != '000000':
            color.set(qn('w:val'), '000000')
    if routing:
        is_pc = value('Approval_Path') == 'Planning Commission Approval'
        threshold = sfr_threshold(context)
        for paragraph in doc.element.iter(qn('w:p')):
            text = ''.join(n.text or '' for n in paragraph.iter(qn('w:t')))
            if text.startswith('BUILDING LINE'):
                _checkboxes(paragraph, [True, False])
            elif text.startswith('Contains lots, at least 60 percent'):
                _checkboxes(paragraph, [threshold is True, threshold is False])
            elif 'SMLS Area Application' in text:
                _checkboxes(paragraph, [False, False, not renewal])
            elif 'SMLS Block Renewal' in text:
                _checkboxes(paragraph, [False])
            elif 'SMBL Block Renewal' in text:
                _checkboxes(paragraph, [renewal])
            elif 'Planning Commission Approval' in text and "Director's Approval" in text:
                _checkboxes(paragraph, [is_pc, not is_pc, True, True, renewal])
    return True
