"""Per-property review values, separate from the area's proposed minimum."""
from decimal import Decimal, InvalidOperation
import re

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
T, P = "{" + W + "}t", "{" + W + "}p"
TOKEN = "«Property_Lot_Size»"
MISSING = "[VERIFY PROPERTY LOT SIZE]"
PREFIX = "Preliminary staff review of the application indicates that your property's lot size is"
# Only the introductory review in primary notifications; the proposed
# restriction and public-hearing notices retain the case minimum.
LEGACY_REVIEW = re.compile(
    r"(?P<prefix>A preliminary review indicates the proposed minimum lot size for this block would be"
    r"|Preliminary staff review of the application indicates that the minimum lot size is)"
    r"\s+(?P<value>[\dXx][\dXx,\s.]*?)\s+(?P<unit>square feet|feet)", re.IGNORECASE)

def property_lot_size(owner):
    """Read this property's Step 1 area, never a proposed or renewal minimum."""
    raw = owner.get("Property_Lot_Size", owner.get("Total_Land_Area"))
    if raw is None or isinstance(raw, bool):
        return MISSING
    text = str(raw).strip().replace(",", "")
    text = re.sub(r"\s*(?:square\s+feet|sq\.?\s*ft\.?|sf)\s*$", "", text, flags=re.I)
    try:
        value = Decimal(text)
    except (InvalidOperation, ValueError):
        return MISSING
    if not value.is_finite() or value <= 0:
        return MISSING
    formatted = format(value, ",f")
    return formatted.rstrip("0").rstrip(".") if "." in formatted else formatted

def _text(paragraph):
    return "".join(node.text or "" for node in paragraph.iter(T)).replace("\u00a0", " ")

def replace_span(paragraph, start, end, replacement):
    """Replace text across Word runs while retaining each run's formatting."""
    offset = 0
    inserted = False
    edited = []
    for node in paragraph.iter(T):
        value = node.text or ""
        limit = offset + len(value)
        if limit > start and offset < end:
            left, right = max(0, start - offset), min(len(value), end - offset)
            node.text = value[:left] + (replacement if not inserted else "") + value[right:]
            node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            inserted = True
            edited.append(node)
        offset = limit
    return edited

def _blacken_value(node):
    """Filled data prints black; the template may retain a red placeholder."""
    from docx.oxml import OxmlElement
    run = node.getparent()
    while run is not None and run.tag != "{" + W + "}r":
        run = run.getparent()
    if run is None:
        return
    properties = run.find("{" + W + "}rPr")
    if properties is None:
        properties = OxmlElement("w:rPr")
        run.insert(0, properties)
    color = properties.find("{" + W + "}color")
    if color is None:
        color = OxmlElement("w:color")
        properties.append(color)
    color.attrib.clear()
    color.set("{" + W + "}val", "000000")
    for tag in ("shd", "highlight"):
        for element in properties.findall("{" + W + "}" + tag):
            properties.remove(element)

def prepare_property_review(root):
    """Give older notification wording an explicit property merge field."""
    changed = 0
    for paragraph in root.iter(P):
        match = LEGACY_REVIEW.search(_text(paragraph))
        if match:
            for group, replacement in (("unit", "square feet"), ("value", TOKEN), ("prefix", PREFIX)):
                replace_span(paragraph, match.start(group), match.end(group), replacement)
            changed += 1
    return changed

def fill_property_review(doc, owner):
    """Fill current and earlier notification templates with this row's area."""
    root = doc.element.body
    prepare_property_review(root)
    area = property_lot_size(owner)
    for paragraph in root.iter(P):
        for marker in (TOKEN, "{{Property_Lot_Size}}"):
            while marker in _text(paragraph):
                start = _text(paragraph).index(marker)
                for node in replace_span(paragraph, start, start + len(marker), area):
                    _blacken_value(node)
