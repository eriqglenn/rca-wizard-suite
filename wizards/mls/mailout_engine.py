from __future__ import annotations

import copy
import re
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

from core.routing_form import prepare_routing_form, routing_pdf_one_page
from docx import Document
from core.docx_unlock import unlock_docx
from core.pdf_condense import condense_hcad_pdfs
from core.sheet_layout import apply_staff_analysis_page_setup
from docx.oxml.ns import qn
from docx.shared import Inches, RGBColor, Pt
from docx.table import Table
from docx.text.paragraph import Paragraph

try:
    import openpyxl
except ImportError as exc:
    raise SystemExit("Install dependencies first: pip install -r requirements.txt") from exc

PLACEHOLDER_RE = re.compile(r"{{\s*([A-Za-z0-9_]+)\s*}}")
STEP1_SHEET = "Step 1 - Input Properties Info"
STEP2_SHEET = "Step 2 - Calculation Sheet"
CASE_INFO_SHEET = "Case Info"

OWNER_FIELD_MAP = {
    "Mailto": "OWNER_MAILTO",
    "Mail_Addr_1": "Mail_Addr_1",
    "Mail_Addr_2": "Mail_Addr_2",
    "Mail_City": "Mail_City",
    "Mail_State": "Mail_State",
    "Mail_Zip": "Mail_Zip",
    "Total_Land_Area": "Total_Land_Area",
    "group_dscr": "group_dscr",
}
# Read when present (the 2026 owner letters print the legal description);
# a Step 1 export without them still loads.
OWNER_OPTIONAL_FIELD_MAP = {
    "Legal_Dscr_1": "Legal_Dscr_1",
    "Legal_Dscr_2": "Legal_Dscr_2",
}

REQUIRED_CASE_FIELDS = [
    "Letter_Date",
    "Protest_Deadline",
    "Planner_Name",
    "Planner_Phone",
    "Planner_Email",
]

OPTIONAL_CASE_FIELDS = [
    "Council_District",
]

REQUIRED_APPLICANT_FIELDS = [
    "Letter_Date",
    "Planner_Name",
    "Planner_Phone",
    "Planner_Email",
    "Applicant_Name",
    "Applicant_Address_1",
    "Applicant_City",
    "Applicant_State",
    "Applicant_Zip",
]

REQUIRED_DIRECTOR_FIELDS = [
    "Director_Name",
    "Neighborhood_Age_Statement",
    "Structure_Age_Statement",
]

REQUIRED_CHECKLIST_FIELDS = [
    "Planner_Name",
    "Planner_Phone",
    "Application_Number",
    "Letter_Date",
    "Approval_Path",
]

REQUIRED_ROUTING_FIELDS = [
    "Planner_Name",
    "Planner_Phone",
    "Application_Number",
    "Council_Deadline",
    "Approval_Path",
]

@dataclass
class WorkbookSummary:
    case_info: Dict[str, str]
    owners: List[Dict[str, str]]
    warnings: List[str]
    subject_line: str


def normalize_value(value: Any) -> str:
    if value is None:
        return ""
    if hasattr(value, "strftime"):
        # Build from parts (no %-d/%#d) so the day has no leading zero and the
        # result is identical on Windows and Linux; a datetime never leaks a
        # '2006-01-18 00:00:00' timestamp into the documents.
        try:
            return f"{value.strftime('%B')} {value.day}, {value.year}"
        except Exception:
            return str(value).strip()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_application_type(value: Any) -> str:
    """Convert worksheet labels like SPECIAL MINIMUM LOT SIZE BLOCK into letter-ready title case.

    The Step 2 calculation sheet uses all-caps labels because it is an official RCA
    calculation page. The mail-out letter should read naturally in the subject line.
    """
    raw = normalize_value(value)
    normalized = re.sub(r"\s+", " ", raw).strip()
    lookup = {
        "SPECIAL MINIMUM LOT SIZE BLOCK": "Special Minimum Lot Size Block Application",
        "SPECIAL MINIMUM LOT SIZE BLOCK RENEWAL": "Special Minimum Lot Size Block Renewal Application",
        "SPECIAL MINIMUM LOT SIZE AREA": "Special Minimum Lot Size Area Application",
        "SPECIAL MINIMUM BUILDING LINE BLOCK": "Special Minimum Building Line Block Application",
    }
    if normalized.upper() in lookup:
        return lookup[normalized.upper()]
    # Fallback: title case, then restore common acronyms/terms.
    text = normalized.title()
    text = text.replace("Smls", "SMLS").replace("Smbl", "SMBL")
    return text


def combine_with_and(left: Any, right: Any, prefix: str = "", suffix: str = "") -> str:
    left_s = normalize_value(left)
    right_s = normalize_value(right)
    if left_s and right_s:
        middle = f"{left_s} and {right_s}"
    else:
        middle = left_s or right_s
    return f"{prefix}{middle}{suffix}".strip()


def read_step2_case_defaults(wb) -> Dict[str, str]:
    if STEP2_SHEET not in wb.sheetnames:
        return {}
    ws = wb[STEP2_SHEET]
    return {
        "Application_Number": normalize_value(ws["B2"].value),
        "Application_Type": normalize_application_type(ws["A1"].value),
        "Street_Name": normalize_value(ws["B4"].value),
        "Lot_Block_Description": normalize_value(ws["D4"].value),
        "Block_Description": normalize_value(ws["D4"].value),
        "Side_Of_Street": combine_with_and(ws["B6"].value, ws["D6"].value, suffix=" sides"),
        "Cross_Streets": combine_with_and(ws["B5"].value, ws["D5"].value, prefix="between "),
        # E4 on Step 2 is often a label like "6250-Square Foot" (a CONCATENATE in
        # the workbook). The owner letter template already supplies the words
        # "square feet" as fixed text, so merge only the clean number here to avoid
        # rendering "6250-Square Foot". lot_size_number extracts and comma-formats
        # the number (e.g. "6250-Square Foot" -> "6,250"); a value that is already
        # a plain number passes through unchanged.
        "Proposed_Min_Lot_Size": lot_size_number(ws["E4"].value),
        "Minimum_Lot_Size": lot_size_number(ws["E4"].value),
    }


def read_case_info(wb) -> Dict[str, str]:
    data: Dict[str, str] = {}
    if CASE_INFO_SHEET not in wb.sheetnames:
        return data
    ws = wb[CASE_INFO_SHEET]
    # Preferred layout: row 4 has Field_Name and Value, data starts on row 5.
    headers = [normalize_value(c.value) for c in ws[4]] if ws.max_row >= 4 else []
    if "Field_Name" in headers and "Value" in headers:
        fi = headers.index("Field_Name") + 1
        vi = headers.index("Value") + 1
        for row in range(5, ws.max_row + 1):
            key = normalize_value(ws.cell(row, fi).value)
            val = normalize_value(ws.cell(row, vi).value)
            if key:
                data[key] = val
        return data
    # Fallback layout: two columns beginning at A1.
    for row in range(1, ws.max_row + 1):
        key = normalize_value(ws.cell(row, 1).value)
        val = normalize_value(ws.cell(row, 2).value)
        if key:
            data[key] = val
    return data


def read_owner_rows_from_step1(wb) -> List[Dict[str, str]]:
    if STEP1_SHEET not in wb.sheetnames:
        raise ValueError(f"Workbook must contain the '{STEP1_SHEET}' tab.")
    ws = wb[STEP1_SHEET]
    headers = [normalize_value(c.value) for c in ws[1]]
    header_to_index = {h: i + 1 for i, h in enumerate(headers) if h}

    missing = [gis_name for gis_name in OWNER_FIELD_MAP.values() if gis_name not in header_to_index]
    if missing:
        raise ValueError(f"Step 1 is missing required GIS columns: {', '.join(missing)}")

    rows: List[Dict[str, str]] = []
    for r in range(2, ws.max_row + 1):
        row: Dict[str, str] = {}
        for template_field, gis_col in OWNER_FIELD_MAP.items():
            row[template_field] = normalize_value(ws.cell(r, header_to_index[gis_col]).value)
        for template_field, gis_col in OWNER_OPTIONAL_FIELD_MAP.items():
            v = normalize_value(ws.cell(r, header_to_index[gis_col]).value) if gis_col in header_to_index else ""
            row[template_field] = "" if v in ("<Null>", "NULL", "None") else v
        if not row.get("Mailto") or not row.get("Mail_Addr_1"):
            continue
        row["Excel_Row"] = str(r)
        row["Mail_Addr_2_Block"] = ("\n" + row["Mail_Addr_2"]) if row.get("Mail_Addr_2") else ""
        rows.append(row)
    from core.lot_divisions import collapse_lot_notifications
    return collapse_lot_notifications(wb,rows)



def collect_hcad_records_from_step1(excel_path):
    """One HCAD property record per normalized account, including separated lots."""
    from core.hcad_identity import collect_hcad_records
    return collect_hcad_records(excel_path)



def collect_hcad_addresses_from_step1(excel_path: str | Path) -> List[str]:
    """Backward-compatible address-only list for older callers/tests."""
    return [r["Address"] for r in collect_hcad_records_from_step1(excel_path) if r.get("Address")]


def export_hcad_address_list(excel_path: str | Path, output_folder: str | Path) -> Tuple[Path, int]:
    """Write Step 1 HCAD search records to Output/HCAD_Records/HCAD_Address_List.txt."""
    case_info, _owners = load_case_and_owners(excel_path)
    app_no = re.sub(r"[^A-Za-z0-9_-]+", "", case_info.get("Application_Number", "")) or "Case"
    hcad_folder = Path(output_folder) / "HCAD_Records"
    hcad_folder.mkdir(parents=True, exist_ok=True)
    address_path = hcad_folder / f"MLS_{app_no}_HCAD_Address_List.txt"
    records = collect_hcad_records_from_step1(excel_path)
    if not records:
        raise ValueError("No site addresses or TAX_ID values were found on Step 1 for HCAD download.")
    lines = []
    for rec in records:
        address = rec.get("Address", "")
        tax_id = rec.get("TAX_ID", "")
        if tax_id:
            lines.append(f"{address}\t{tax_id}")
        else:
            lines.append(address)
    address_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return address_path, len(records)

def build_subject_line(case_info: Dict[str, str]) -> str:
    return "\n".join([
        case_info.get("Block_Description") or case_info.get("Lot_Block_Description", ""),
        case_info.get("Side_Of_Street", ""),
        case_info.get("Cross_Streets", ""),
    ]).strip()


def validate_inputs(case_info: Dict[str, str], owner_rows: List[Dict[str, str]], map_path: Path | None = None) -> List[str]:
    warnings: List[str] = []
    for field in REQUIRED_CASE_FIELDS:
        if not case_info.get(field):
            warnings.append(f"Missing Case Info field: {field}")
    if not owner_rows:
        warnings.append("No valid owner rows found on Step 1.")
    for i, owner in enumerate(owner_rows, start=1):
        for field in ["Mailto", "Mail_Addr_1", "Mail_City", "Mail_State", "Mail_Zip", "Total_Land_Area", "group_dscr"]:
            if not owner.get(field):
                warnings.append(f"Owner #{i} / Excel row {owner.get('Excel_Row', '?')} is missing {field}.")
    if map_path is not None and not map_path.exists():
        warnings.append(f"Map file not found: {map_path}")
    return warnings


def load_case_and_owners(excel_path: str | Path) -> Tuple[Dict[str, str], List[Dict[str, str]]]:
    excel_path = Path(excel_path)
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    # Planner-entered Case Info remains editable, but the blue Auto-Read Case Values
    # must come from Step 2 directly. This avoids stale Excel formula caches and
    # workbook-side formula errors such as #NAME? from overriding the source of truth.
    step2_defaults = read_step2_case_defaults(wb)
    case_info = read_case_info(wb)
    for key, value in step2_defaults.items():
        if value and str(value).strip().upper() not in {"#NAME?", "#VALUE!", "#REF!", "#N/A", "#DIV/0!", "#NULL!", "#NUM!"}:
            case_info[key] = value

    # Case Info may contain formulas that mirror the official all-caps Step 2 label.
    # Always normalize application type for public-facing letters after all workbook values are merged.
    if case_info.get("Application_Type"):
        case_info["Application_Type"] = normalize_application_type(case_info["Application_Type"])
    if not case_info.get("Block_Description"):
        case_info["Block_Description"] = case_info.get("Lot_Block_Description", "")
    if not case_info.get("Minimum_Lot_Size"):
        case_info["Minimum_Lot_Size"] = case_info.get("Proposed_Min_Lot_Size", "")
    # Applicant-letter convenience fields and defaults. Address line 2 is suppressed when blank.
    addr2 = case_info.get("Applicant_Address_2", "")
    case_info["Applicant_Address_2_Block"] = ("\n" + addr2) if addr2 else ""
    case_info.setdefault("Sign_Count_Number", "2")
    case_info.setdefault("Sign_Count_Words", "two")
    case_info.setdefault("Sign_Placement", "one sign at each corner")
    if not case_info.get("Approval_Path"):
        case_info["Approval_Path"] = "Director Approval"
    owners = read_owner_rows_from_step1(wb)
    return case_info, owners


def summarize_workbook(excel_path: str | Path, map_path: str | Path | None = None) -> WorkbookSummary:
    case_info, owners = load_case_and_owners(excel_path)
    resolved_map = Path(map_path) if map_path else None
    warnings = validate_inputs(case_info, owners, resolved_map)
    return WorkbookSummary(case_info=case_info, owners=owners, warnings=warnings, subject_line=build_subject_line(case_info))


def replace_text_in_paragraph(paragraph: Paragraph, values: Dict[str, str]) -> None:
    """Replace placeholders without flattening the paragraph.

    This preserves the City template's run-level formatting (bold, font, size) instead
    of collapsing the entire paragraph into the first run. Placeholders should be
    placed inside single runs in the template.
    """
    if not paragraph.runs or "{{MAP_IMAGE}}" in paragraph.text or "{{" not in paragraph.text:
        return

    def repl(match):
        key = match.group(1)
        return values.get(key, match.group(0))

    for run in paragraph.runs:
        if "{{" in run.text:
            new_text = PLACEHOLDER_RE.sub(repl, run.text)
            if new_text != run.text:
                run.text = new_text
                # Filled values should print as normal document text while keeping the
                # template's font, size, bold/italic, and spacing.
                run.font.color.rgb = RGBColor(0, 0, 0)


def iter_wrapped_paragraphs_and_tables(doc: Document, elements: Iterable[Any] | None = None):
    parent = doc
    source = elements if elements is not None else doc.element.body.iterchildren()
    for element in source:
        if element.tag == qn("w:p"):
            yield Paragraph(element, parent)
        elif element.tag == qn("w:tbl"):
            table = Table(element, parent)
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        yield paragraph


def convert_pdf_first_page_to_png(map_path: Path) -> Path:
    """Return an image path suitable for insertion into Word.

    v13.1 removes the Poppler-only dependency. If the selected map is a PDF,
    the wizard first renders page 1 with PyMuPDF, which installs cleanly from
    pip on Windows. Poppler/pdftoppm remains as a fallback only.
    """
    if map_path.suffix.lower() != ".pdf":
        return map_path

    outdir = Path(tempfile.mkdtemp(prefix="mailout_map_"))
    outpng = outdir / "map.png"

    # Preferred path: PyMuPDF. This avoids requiring planners to install Poppler.
    try:
        try:
            import fitz  # PyMuPDF
        except ImportError:
            import sys
            subprocess.check_call([sys.executable, "-m", "pip", "install", "PyMuPDF>=1.24.0"])
            import fitz  # type: ignore

        doc = fitz.open(str(map_path))
        if doc.page_count == 0:
            raise ValueError(f"PDF map has no pages: {map_path}")
        page = doc.load_page(0)
        # 2x scale gives a crisp letter-sized map without oversized files.
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        pix.save(str(outpng))
        doc.close()
        return outpng
    except Exception:
        # Fallback path for machines that already have Poppler installed.
        pdftoppm = shutil.which("pdftoppm")
        if pdftoppm:
            outstem = outdir / "map"
            subprocess.run([pdftoppm, "-png", "-singlefile", "-r", "150", str(map_path), str(outstem)], check=True)
            return outstem.with_suffix(".png")

        raise ValueError(
            "The wizard could not convert the selected PDF map. "
            "Run install_requirements.bat so PyMuPDF is installed, or use a PNG/JPG map."
        )


def fill_elements(doc: Document, elements: Iterable[Any] | None, values: Dict[str, str], map_image_path: Path) -> None:
    for paragraph in iter_wrapped_paragraphs_and_tables(doc, elements):
        if "{{MAP_IMAGE}}" in paragraph.text:
            for run in paragraph.runs:
                run.text = ""
            run = paragraph.add_run()
            # Width chosen to fit within 1-inch Word margins on letter paper.
            run.add_picture(str(map_image_path), width=Inches(6.35))
            paragraph.alignment = 1
        else:
            replace_text_in_paragraph(paragraph, values)


def _strip_section_properties(el) -> None:
    # Avoid copied section breaks, which are the most common cause of blank pages
    # when DOCX templates are duplicated programmatically.
    for sect in list(el.iter()):
        if sect.tag.endswith("sectPr"):
            parent = sect.getparent()
            if parent is not None:
                parent.remove(sect)


def get_template_body_elements(template_path: Path) -> List[Any]:
    template = Document(template_path)
    elements: List[Any] = []
    for el in template.element.body:
        if el.tag.endswith("sectPr"):
            continue
        new_el = copy.deepcopy(el)
        _strip_section_properties(new_el)
        elements.append(new_el)
    return elements


def _body_insert_before_sectpr(body, new_el) -> None:
    # Word expects normal content before the final w:sectPr. Appending after it can
    # create invalid ordering and render as blank pages.
    sectpr = None
    for child in body:
        if child.tag.endswith("sectPr"):
            sectpr = child
            break
    if sectpr is not None:
        body.insert(body.index(sectpr), new_el)
    else:
        body.append(new_el)


def append_template_copy(master: Document, template_elements: List[Any]) -> List[Any]:
    body = master.element.body
    inserted: List[Any] = []
    for el in template_elements:
        new_el = copy.deepcopy(el)
        _strip_section_properties(new_el)
        _body_insert_before_sectpr(body, new_el)
        inserted.append(new_el)
    return inserted



NUMBER_WORDS = {
    0: "zero", 1: "one", 2: "two", 3: "three", 4: "four", 5: "five", 6: "six", 7: "seven", 8: "eight", 9: "nine",
    10: "ten", 11: "eleven", 12: "twelve", 13: "thirteen", 14: "fourteen", 15: "fifteen", 16: "sixteen",
    17: "seventeen", 18: "eighteen", 19: "nineteen", 20: "twenty", 21: "twenty-one", 22: "twenty-two",
    23: "twenty-three", 24: "twenty-four", 25: "twenty-five", 26: "twenty-six", 27: "twenty-seven", 28: "twenty-eight",
    29: "twenty-nine", 30: "thirty", 31: "thirty-one", 32: "thirty-two", 33: "thirty-three", 34: "thirty-four",
    35: "thirty-five", 36: "thirty-six", 37: "thirty-seven", 38: "thirty-eight", 39: "thirty-nine", 40: "forty",
    41: "forty-one", 42: "forty-two", 43: "forty-three", 44: "forty-four", 45: "forty-five", 46: "forty-six",
    47: "forty-seven", 48: "forty-eight", 49: "forty-nine", 50: "fifty",
}

def number_word(n: int, capitalize: bool = False) -> str:
    word = NUMBER_WORDS.get(n, str(n))
    return word[:1].upper() + word[1:] if capitalize and word else word


def pct(value: Any) -> str:
    if value is None or value == "":
        return ""
    try:
        v = float(value)
        if abs(v) <= 1:
            v *= 100
        return str(int(round(v)))
    except Exception:
        raw = normalize_value(value).replace("%", "")
        return raw


def lot_size_number(value: Any) -> str:
    raw = normalize_value(value)
    m = re.search(r"[0-9][0-9,]*", raw)
    if not m:
        return raw
    n = int(m.group(0).replace(",", ""))
    return f"{n:,}"


def read_director_defaults_from_step2(excel_path: str | Path) -> Dict[str, str]:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if STEP2_SHEET not in wb.sheetnames:
        return {}
    ws = wb[STEP2_SHEET]
    defaults: Dict[str, str] = {}
    defaults["Support_Percentage"] = pct(ws["F8"].value if ws["F8"].value not in (None, "") else ws["H8"].value)
    defaults["SFR_Percentage"] = pct(ws["H11"].value)
    total_lots = None
    for cell in ["A15", "F11"]:
        try:
            v = ws[cell].value
            if v not in (None, ""):
                total_lots = int(float(v))
                break
        except Exception:
            pass
    defaults["Total_Lot_Count"] = normalize_value(total_lots)
    min_lot = lot_size_number(ws["E4"].value)
    defaults["Minimum_Lot_Size_Number"] = min_lot

    qualifying_count = None
    # Planner-entered count of lots meeting the minimum lot size, read from the
    # Case Info field Lots_Meeting_Minimum (the wizard does not calculate it).
    try:
        _ci = openpyxl.load_workbook(excel_path, data_only=True)
        if CASE_INFO_SHEET in _ci.sheetnames:
            _w = _ci[CASE_INFO_SHEET]
            for _r in range(1, _w.max_row + 1):
                if _w.cell(_r, 1).value and str(_w.cell(_r, 1).value).strip() == "Lots_Meeting_Minimum":
                    _v = _w.cell(_r, 2).value
                    if _v not in (None, ""):
                        qualifying_count = int(float(str(_v).replace(",", "")))
                    break
    except Exception:
        qualifying_count = None
    if qualifying_count is None:
        # Blank -> leave unset so the document keeps its visible placeholder.
        defaults["Qualifying_Lot_Count"] = ""
        defaults["Qualifying_Lot_Count_Words"] = ""
        defaults["Qualifying_Lot_Count_Words_Cap"] = ""
        defaults["Total_Lot_Count_Words"] = number_word(total_lots) if total_lots is not None else ""
        defaults["Qualifying_Area_Percentage"] = ""
        return defaults
    qualifying_area = None
    defaults["Qualifying_Lot_Count"] = str(qualifying_count)
    defaults["Qualifying_Lot_Count_Words"] = number_word(qualifying_count)
    defaults["Qualifying_Lot_Count_Words_Cap"] = number_word(qualifying_count, capitalize=True)
    if total_lots is not None:
        defaults["Total_Lot_Count_Words"] = number_word(total_lots)
    else:
        defaults["Total_Lot_Count_Words"] = ""
    defaults["Qualifying_Area_Percentage"] = pct(qualifying_area) if qualifying_area is not None else ""
    return defaults


def build_director_context(excel_path: str | Path) -> Dict[str, str]:
    case_info, _owner_rows = load_case_and_owners(excel_path)
    case_info.update(read_director_defaults_from_step2(excel_path))
    director_block = ", ".join([x for x in [case_info.get("Block_Description", ""), case_info.get("Side_Of_Street", "")] if x])
    cross = case_info.get("Cross_Streets", "")
    if cross:
        director_block = f"{director_block} {cross}".strip()
    case_info["Director_Block_Description"] = director_block
    if not case_info.get("Director_Name"):
        case_info["Director_Name"] = "Vonn Tran"
    if not case_info.get("Director_Approval_Date"):
        case_info["Director_Approval_Date"] = ""
    if not case_info.get("Preservation_Finding"):
        mls = case_info.get("Minimum_Lot_Size_Number") or lot_size_number(case_info.get("Minimum_Lot_Size"))
        case_info["Preservation_Finding"] = f"The establishment of a {mls} square foot minimum lot size will preserve the lot size character of the area." if mls else ""
    return case_info


def validate_director_inputs(case_info: Dict[str, str]) -> List[str]:
    warnings: List[str] = []
    for field in REQUIRED_DIRECTOR_FIELDS:
        if not case_info.get(field):
            warnings.append(f"Missing Case Info Director Approval field: {field}")
    for field in ["Application_Number", "Director_Block_Description", "Support_Percentage", "SFR_Percentage", "Minimum_Lot_Size_Number", "Total_Lot_Count"]:
        if not case_info.get(field):
            warnings.append(f"Missing calculation field for Director Approval: {field}")
    return warnings


def generate_director_approval(
    excel_path: str | Path,
    template_path: str | Path,
    output_path: str | Path,
) -> int:
    from core.city_templates import render_updated
    from wizards.mls.templates_2026 import build_context
    if render_updated(template_path, output_path, build_context(excel_path), excel_path, 'mls', 'director'): return 1
    excel_path = Path(excel_path)
    template_path = Path(template_path)
    output_path = Path(output_path)
    context = build_director_context(excel_path)
    warnings = validate_director_inputs(context)
    if warnings:
        raise ValueError("Cannot generate Director Approval until these issues are fixed:\n" + "\n".join(warnings))
    doc = Document(template_path)
    fill_elements(doc, None, context, Path("__no_map_needed__.png"))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    unlock_docx(output_path)  # fully editable output
    return 1


def _apply_text_replacements_preserving_first_run(doc: Document, replacements: Dict[str, str]) -> None:
    """Replace text in paragraphs/cells while preserving the paragraph's existing style.

    This is used for fixed City templates where placeholders may not exist. For a
    paragraph that needs a cross-run replacement (for example 832-393-XXXX split
    across runs), the visible paragraph text is rewritten into the first run and the
    paragraph formatting remains from the template. This is limited to fields whose
    runs share the same effective style in the checklist template.
    """
    for paragraph in iter_wrapped_paragraphs_and_tables(doc):
        text = paragraph.text
        new_text = text
        for old, new in replacements.items():
            if old in new_text:
                new_text = new_text.replace(old, new)
        if new_text != text:
            if paragraph.runs:
                paragraph.runs[0].text = new_text
                paragraph.runs[0].font.color.rgb = RGBColor(0, 0, 0)
                for run in paragraph.runs[1:]:
                    run.text = ""


def _replace_first_run_text(doc: Document, old: str, new: str) -> None:
    for paragraph in iter_wrapped_paragraphs_and_tables(doc):
        for run in paragraph.runs:
            if old in run.text:
                run.text = run.text.replace(old, new)
                run.font.color.rgb = RGBColor(0, 0, 0)
                return


def _checkbox_state(checked: bool) -> str:
    return "☒" if checked else "☐"


def patch_docx_checkboxes(docx_path: str | Path, label_to_checked: Dict[str, bool]) -> None:
    """Patch Word checkbox content controls next to a given label.

    The official RCA checklist uses Word content-control checkboxes. python-docx
    does not expose those controls, so this patches the underlying OOXML while
    preserving the City's layout, fonts, spacing, and table structure.
    """
    from lxml import etree

    docx_path = Path(docx_path)
    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    }
    with zipfile.ZipFile(docx_path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}
    root = etree.fromstring(files["word/document.xml"])

    def set_sdt(sdt, checked: bool) -> None:
        checked_el = sdt.xpath(".//w14:checked", namespaces=ns)
        if checked_el:
            checked_el[0].set(f"{{{ns['w14']}}}val", "1" if checked else "0")
        text_els = sdt.xpath(".//w:sdtContent//w:t", namespaces=ns)
        if text_els:
            text_els[0].text = _checkbox_state(checked)

    for p in root.xpath(".//w:p", namespaces=ns):
        paragraph_text = "".join(p.xpath(".//w:t/text()", namespaces=ns))
        sdts = p.xpath(".//w:sdt[w:sdtPr/w14:checkbox]", namespaces=ns)
        if not sdts:
            continue

        # Some RCA forms place multiple checkbox controls in one paragraph.
        # Map them by order so later checkboxes do not retain the template defaults.
        if "Building Line" in paragraph_text and "Lot Size" in paragraph_text and len(sdts) >= 2:
            set_sdt(sdts[0], label_to_checked.get("Building Line", False))
            set_sdt(sdts[1], label_to_checked.get("Lot Size", False))
            continue
        if "SMLS Area Application" in paragraph_text and "SMLS Block Application" in paragraph_text and "SMBL Block Application" in paragraph_text and len(sdts) >= 3:
            set_sdt(sdts[0], label_to_checked.get("SMLS Area Application", False))
            set_sdt(sdts[1], label_to_checked.get("SMLS Block Application", False))
            set_sdt(sdts[2], label_to_checked.get("SMBL Block Application", False))
            continue
        if "SMBL Block Application" in paragraph_text and "SMLS Block Renewal" in paragraph_text and len(sdts) >= 2:
            set_sdt(sdts[0], label_to_checked.get("SMBL Block Application", False))
            set_sdt(sdts[1], label_to_checked.get("SMLS Block Renewal", False))
            continue
        if "SMLS Block Renewal" in paragraph_text and "SMBL Block Renewal" in paragraph_text and len(sdts) >= 2:
            set_sdt(sdts[0], label_to_checked.get("SMLS Block Renewal", False))
            set_sdt(sdts[1], label_to_checked.get("SMBL Block Renewal", False))
            continue
        if "Planning Commission Approval" in paragraph_text and "Original Ordinance" in paragraph_text and len(sdts) >= 8:
            order = [
                "Planning Commission Approval", "Director's Approval", "Map of the Area", "Staff Analysis",
                "Copy of Subdivision Plat", "Tax Map", "Tax Records", "Original Ordinance"
            ]
            for sdt, label in zip(sdts, order):
                set_sdt(sdt, label_to_checked.get(label, False))
            continue

        for label, checked in label_to_checked.items():
            if label in paragraph_text:
                set_sdt(sdts[0], checked)
                break

    files["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    tmp = docx_path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    tmp.replace(docx_path)



def remove_final_empty_paragraph(docx_path: str | Path) -> None:
    """Remove a trailing empty paragraph before the final section properties.

    Some City templates contain a final blank paragraph that LibreOffice/Word can
    render as a trailing blank page. Removing only that final empty paragraph keeps
    the visible template content intact while preventing blank output pages.
    """
    from lxml import etree
    docx_path = Path(docx_path)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with zipfile.ZipFile(docx_path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}
    root = etree.fromstring(files["word/document.xml"])
    body = root.find("w:body", ns)
    if body is not None:
        children = list(body)
        if len(children) >= 2 and children[-1].tag.endswith("sectPr") and children[-2].tag.endswith("p"):
            text = "".join(children[-2].xpath(".//w:t/text()", namespaces=ns)).replace("\u00a0", "").strip()
            if not text:
                body.remove(children[-2])
    files["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    tmp = docx_path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    tmp.replace(docx_path)



def remove_blank_body_paragraphs(docx_path: str | Path) -> None:
    """Remove empty body paragraphs that can create a trailing blank page after routing-form fills."""
    from lxml import etree
    docx_path = Path(docx_path)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with zipfile.ZipFile(docx_path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}
    root = etree.fromstring(files["word/document.xml"])
    body = root.find("w:body", ns)
    if body is not None:
        for p in list(body.findall("w:p", ns)):
            text = "".join(p.xpath(".//w:t/text()", namespaces=ns)).replace("\u00a0", "").strip()
            has_visual = bool(p.xpath(".//w:drawing|.//w:pict|.//w:object|.//w:br", namespaces=ns))
            if not text and not has_visual:
                body.remove(p)
    files["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    tmp = docx_path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    tmp.replace(docx_path)

def remove_red_placeholder_text_formatting(docx_path: str | Path) -> None:
    """Remove red placeholder text color from filled fields, while preserving red separator lines.

    The City template uses red text for editable placeholders and also red separator
    rules made from underscore characters. Filled values should be black, but the
    template's red separator lines should remain red.
    """
    from lxml import etree
    docx_path = Path(docx_path)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with zipfile.ZipFile(docx_path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}
    root = etree.fromstring(files["word/document.xml"])
    for p in root.xpath(".//w:p", namespaces=ns):
        ptext = "".join(p.xpath(".//w:t/text()", namespaces=ns))
        # Preserve the red horizontal separator rules, which are underscore-only paragraphs.
        if ptext and not ptext.replace("_", "").strip():
            continue
        for rpr in p.xpath(".//w:rPr", namespaces=ns):
            for child in list(rpr):
                local = child.tag.split('}')[-1]
                if local == "color":
                    rpr.remove(child)
    files["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    tmp = docx_path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    tmp.replace(docx_path)

def build_checklist_context(excel_path: str | Path) -> Dict[str, str]:
    case_info, _owners = load_case_and_owners(excel_path)
    # This wizard handles standard MLS applications only, not renewals.
    if not case_info.get("Approval_Path"):
        case_info["Approval_Path"] = "Director Approval"
    return case_info


def validate_checklist_inputs(case_info: Dict[str, str]) -> List[str]:
    warnings: List[str] = []
    for field in REQUIRED_CHECKLIST_FIELDS:
        if not case_info.get(field):
            warnings.append(f"Missing Case Info/RCA Checklist field: {field}")
    if case_info.get("Approval_Path") not in ("Director Approval", "Planning Commission Approval"):
        warnings.append("Approval_Path must be either 'Director Approval' or 'Planning Commission Approval'.")
    return warnings


def generate_rca_checklist(
    excel_path: str | Path,
    template_path: str | Path,
    output_path: str | Path,
) -> int:
    excel_path = Path(excel_path)
    template_path = Path(template_path)
    output_path = Path(output_path)
    context = build_checklist_context(excel_path)
    warnings = validate_checklist_inputs(context)
    if warnings:
        raise ValueError("Cannot generate RCA Checklist until these issues are fixed:\n" + "\n".join(warnings))

    doc = Document(template_path)
    app_no = context.get("Application_Number", "")
    planner = context.get("Planner_Name", "")
    phone = context.get("Planner_Phone", "")
    date = context.get("Letter_Date", "")

    replacements = {
        "Planner Name": planner,
        "832-393-XXXX": phone,
        "Date:                       Prepared by:": f"Date: {date}            Prepared by:",
    }
    _apply_text_replacements_preserving_first_run(doc, replacements)
    _replace_first_run_text(doc, "______", str(app_no))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)

    approval_path = context.get("Approval_Path", "Director Approval")
    checkbox_map = {
        "Building Line": False,
        "Lot Size": True,
        "SMLS Area Application": False,
        "SMLS Block Application": True,
        "SMBL Block Application": False,
        "SMLS Block Renewal": False,
        "Planning Commission Approval": approval_path == "Planning Commission Approval",
        "Director's Approval": approval_path == "Director Approval",
        "Map of the Area": True,
        "Staff Analysis": True,
        "Copy of Subdivision Plat": True,
        "Tax Map": True,
        "Tax Records": True,
        "Original Ordinance": False,
    }
    patch_docx_checkboxes(output_path, checkbox_map)
    remove_red_placeholder_text_formatting(output_path)
    remove_final_empty_paragraph(output_path)
    unlock_docx(output_path)  # fully editable output
    return 1


def _set_direct_sdt_texts(paragraph_el, texts: List[str]) -> None:
    """Set text on direct child content controls in order without flattening the paragraph."""
    ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    direct_sdts = [child for child in list(paragraph_el) if child.tag == f"{{{ns_w}}}sdt"]
    for sdt, new_text in zip(direct_sdts, texts):
        text_els = sdt.xpath(".//w:sdtContent//w:t", namespaces={"w": ns_w})
        if not text_els:
            continue
        text_els[0].text = new_text
        for t in text_els[1:]:
            t.text = ""
        # Remove red placeholder color/underlines from visible content controls.
        for rpr in sdt.xpath(".//w:sdtContent//w:rPr", namespaces={"w": ns_w}):
            for child in list(rpr):
                local = child.tag.split('}')[-1]
                if local in ("color", "u"):
                    rpr.remove(child)


def _patch_sdt_text_by_exact(docx_path: str | Path, exact_to_new: Dict[str, str]) -> None:
    """Patch non-checkbox content controls by their exact visible placeholder text."""
    from lxml import etree
    docx_path = Path(docx_path)
    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    }
    with zipfile.ZipFile(docx_path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}
    root = etree.fromstring(files["word/document.xml"])
    for sdt in root.xpath(".//w:sdt", namespaces=ns):
        # Skip checkbox SDTs; they are handled separately.
        if sdt.xpath("./w:sdtPr/w14:checkbox", namespaces=ns):
            continue
        text = "".join(sdt.xpath(".//w:sdtContent//w:t/text()", namespaces=ns))
        if text in exact_to_new:
            text_els = sdt.xpath(".//w:sdtContent//w:t", namespaces=ns)
            if text_els:
                text_els[0].text = exact_to_new[text]
                for t in text_els[1:]:
                    t.text = ""
                for rpr in sdt.xpath(".//w:sdtContent//w:rPr", namespaces=ns):
                    for child in list(rpr):
                        local = child.tag.split('}')[-1]
                        if local in ("color", "u"):
                            rpr.remove(child)
    files["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    tmp = docx_path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    tmp.replace(docx_path)


def patch_routing_form_fields(docx_path: str | Path, context: Dict[str, str]) -> None:
    """Fill the exact MLS Routing Form template while preserving its layout."""
    from lxml import etree
    docx_path = Path(docx_path)
    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    }
    with zipfile.ZipFile(docx_path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}
    root = etree.fromstring(files["word/document.xml"])

    topic = f"Special Minimum Lot Size Block Application No. {context.get('Application_Number','')}"
    location = f"{context.get('Block_Description','')}, {context.get('Side_Of_Street','')}, {context.get('Cross_Streets','')}".replace(" ,", ",").strip().strip(',')
    if location and not location.lower().startswith("the "):
        location = f"The {location}"
    author = ", ".join([v for v in [context.get("Planner_Name", ""), context.get("Planner_Phone", "")] if v])
    council_deadline = context.get("Council_Deadline", "")

    # Top content controls in the routing form.
    replacements = {
        "Special Minimum Lot Size/ Block Application No. X": topic,
        "Special Minimum Lot Size Block Application No. X": topic,
        "The X block of X Street, X and X sides, between X Street and X Street": location,
        "Planner Name, Phone Number": author,
        "June 11, 2026": council_deadline,
    }
    for sdt in root.xpath(".//w:sdt", namespaces=ns):
        if sdt.xpath("./w:sdtPr/w14:checkbox", namespaces=ns):
            continue
        text = "".join(sdt.xpath(".//w:sdtContent//w:t/text()", namespaces=ns))
        if text in replacements:
            text_els = sdt.xpath(".//w:sdtContent//w:t", namespaces=ns)
            if text_els:
                text_els[0].text = replacements[text]
                for t in text_els[1:]:
                    t.text = ""
                # Remove red placeholder color/underline from filled content controls.
                for rpr in sdt.xpath(".//w:sdtContent//w:rPr", namespaces=ns):
                    for child in list(rpr):
                        local = child.tag.split('}')[-1]
                        if local in ("color", "u"):
                            rpr.remove(child)

    # Fix the location line indent so wrapped street names align with the rest of the
    # parenthetical location text instead of drifting left under the label.
    for p in root.xpath(".//w:p", namespaces=ns):
        ptext = "".join(p.xpath(".//w:t/text()", namespaces=ns))
        if location and location in ptext:
            ppr = p.find("w:pPr", namespaces=ns)
            if ppr is None:
                ppr = etree.Element(f"{{{ns['w']}}}pPr")
                p.insert(0, ppr)
            ind = ppr.find("w:ind", namespaces=ns)
            if ind is None:
                ind = etree.SubElement(ppr, f"{{{ns['w']}}}ind")
            # Original template used left=1440 + firstLine=720, which causes wrapped
            # lines to start 0.5 inch left of the first line. Use a normal left indent
            # at the same visual start point and no first-line indent.
            ind.set(f"{{{ns['w']}}}left", "2160")
            ind.set(f"{{{ns['w']}}}firstLine", "0")
            for rpr in p.xpath(".//w:rPr", namespaces=ns):
                for child in list(rpr):
                    local = child.tag.split('}')[-1]
                    if local in ("color", "u"):
                        rpr.remove(child)

    # Legal Department Routing Form detail line. Patch direct SDTs by paragraph.
    for p in root.xpath(".//w:p", namespaces=ns):
        ptext = "".join(p.xpath(".//w:t/text()", namespaces=ns))
        if "BUILDING LINE" in ptext and "LOT SIZE" in ptext and "SFR" in ptext:
            direct_sdts = [child for child in list(p) if child.tag == f"{{{ns['w']}}}sdt"]
            # Expected direct SDTs: building checkbox, lot checkbox, application #, lot size, SFR.
            if len(direct_sdts) >= 5:
                # application number
                _set_direct_sdt_texts(p, ["☐", "☒", context.get("Application_Number", ""), f": {context.get('Minimum_Lot_Size_Number','')}", context.get("SFR_Percentage", "")])
        elif "Petition Signed By" in ptext and "Total Lots" in ptext:
            total_vs_qual = f"{context.get('Total_Lot_Count','')} / {context.get('Qualifying_Lot_Count','')}"
            _set_direct_sdt_texts(p, [f" {context.get('Support_Percentage','')}", total_vs_qual])

    files["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    tmp = docx_path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    tmp.replace(docx_path)


def build_routing_context(excel_path: str | Path) -> Dict[str, str]:
    case_info, _owners = load_case_and_owners(excel_path)
    case_info.update(read_director_defaults_from_step2(excel_path))
    if not case_info.get("Approval_Path"):
        case_info["Approval_Path"] = "Director Approval"
    if not case_info.get("Council_Deadline"):
        case_info["Council_Deadline"] = ""
    return case_info


def validate_routing_inputs(case_info: Dict[str, str]) -> List[str]:
    warnings: List[str] = []
    for field in REQUIRED_ROUTING_FIELDS:
        if not case_info.get(field):
            warnings.append(f"Missing Case Info/Routing Form field: {field}")
    for field in ["Block_Description", "Side_Of_Street", "Cross_Streets", "Minimum_Lot_Size_Number", "SFR_Percentage", "Support_Percentage", "Total_Lot_Count"]:
        if not case_info.get(field):
            warnings.append(f"Missing calculation field for Routing Form: {field}")
    if case_info.get("Approval_Path") not in ("Director Approval", "Planning Commission Approval"):
        warnings.append("Approval_Path must be either 'Director Approval' or 'Planning Commission Approval'.")
    return warnings


def generate_routing_form(
    excel_path: str | Path,
    template_path: str | Path,
    output_path: str | Path,
) -> int:
    from core.city_templates import render_updated
    from wizards.mls.templates_2026 import build_context
    if render_updated(template_path, output_path, build_context(excel_path), excel_path, 'mls', 'routing'): return 1
    excel_path = Path(excel_path)
    template_path = Path(template_path)
    output_path = Path(output_path)
    context = build_routing_context(excel_path)
    warnings = validate_routing_inputs(context)
    if warnings:
        raise ValueError("Cannot generate Routing Form until these issues are fixed:\n" + "\n".join(warnings))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(template_path, output_path)
    patch_routing_form_fields(output_path, context)
    approval_path = context.get("Approval_Path", "Director Approval")
    checkbox_map = {
        "Building Line": False,
        "Lot Size": True,
        "SMLS Area Application": False,
        "SMLS Block Application": True,
        "SMBL Block Application": False,
        "SMLS Block Renewal": False,
        "SMBL Block Renewal": False,
        "Planning Commission Approval": approval_path == "Planning Commission Approval",
        "Director's Approval": approval_path == "Director Approval",
        "Map of the Area": True,
        "Staff Analysis": True,
        "Copy of Subdivision Plat": True,
        "Tax Map": True,
        "Tax Records": True,
        "Original Ordinance": False,
    }
    patch_docx_checkboxes(output_path, checkbox_map)
    remove_red_placeholder_text_formatting(output_path)
    remove_blank_body_paragraphs(output_path)
    remove_final_empty_paragraph(output_path)
    doc = Document(output_path)
    prepare_routing_form(doc)
    doc.save(output_path)
    unlock_docx(output_path)  # fully editable output
    return 1

def validate_applicant_inputs(case_info: Dict[str, str]) -> List[str]:
    warnings: List[str] = []
    for field in REQUIRED_APPLICANT_FIELDS:
        if not case_info.get(field):
            warnings.append(f"Missing Case Info applicant field: {field}")
    if not case_info.get("Minimum_Lot_Size"):
        warnings.append("Missing Step 2/Case Info field: Minimum_Lot_Size or Proposed_Min_Lot_Size")
    if not case_info.get("Application_Type"):
        warnings.append("Missing Step 2/Case Info field: Application_Type")
    return warnings


def generate_applicant_letter(
    excel_path: str | Path,
    template_path: str | Path,
    output_path: str | Path,
) -> int:
    excel_path = Path(excel_path)
    template_path = Path(template_path)
    output_path = Path(output_path)
    case_info, _owner_rows = load_case_and_owners(excel_path)
    warnings = validate_applicant_inputs(case_info)
    if warnings:
        raise ValueError("Cannot generate applicant letter until these issues are fixed:\n" + "\n".join(warnings))
    doc = Document(template_path)
    fill_elements(doc, None, case_info, Path("__no_map_needed__.png"))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    unlock_docx(output_path)  # fully editable output
    return 1


def resolve_map_path(cli_map: str | Path | None, excel_path: Path, case_info: Dict[str, str]) -> Path:
    # v4 design: the map belongs to the wizard UI/file picker, not the calculation workbook.
    # Case Info no longer stores Map_File_Name.
    if not cli_map:
        raise ValueError("Select a map file in the wizard before generating letters.")
    return Path(cli_map)


_WT = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _render_single_owner_letter(template_path: Path, values: Dict[str, str], map_image_path: Path) -> Document:
    """Render exactly one owner letter from a FRESH copy of the official template.

    This uses the same code path as the Preview One Letter output (fill_elements),
    so every letter keeps the template's exact section structure: the narrow
    letterhead band (section 0, which carries the first-page council-member
    footer) and the normal-margin body (section 1). Nothing here edits margins,
    page size, section properties, footer references, paragraph spacing, or fonts.
    """
    doc = Document(str(template_path))
    fill_elements(doc, None, values, map_image_path)
    return doc


def _final_sectpr(body):
    last = body[-1] if len(body) else None
    return last if last is not None and last.tag == _WT + "sectPr" else None


def _append_letter_as_own_section(master: Document, letter: Document) -> None:
    """Append one rendered letter to master as its own self-contained section(s).

    The owner-notification template stores the council-member footer as a
    first-page footer (w:titlePg + footerReference type="first") on its first
    section. A true mail merge must give EVERY owner letter its own copy of that
    section structure so the footer and the normal body margins repeat on every
    letter -- not just the first. docxcompose collapses these to a single
    footer reference, which is why later letters lost both their footer and
    their normal margins; this routine instead preserves each letter's own
    sectPr blocks (and thus its footer references) verbatim.

    All letters in one case share the identical council-member footer, so they
    correctly reference the single template footer part already present in the
    master document.
    """
    mbody = master.element.body
    lbody = letter.element.body

    # Close the master's current trailing content into its own section by tucking
    # the master's existing final sectPr into a new section-break paragraph.
    msect = _final_sectpr(mbody)
    sect_break_p = master.element.makeelement(_WT + "p", {})
    ppr = master.element.makeelement(_WT + "pPr", {})
    sect_break_p.append(ppr)
    if msect is not None:
        mbody.remove(msect)
        ppr.append(msect)
    mbody.append(sect_break_p)

    # Append the next letter's body content, preserving any intermediate sectPr
    # (the letterhead-band -> body break) so its margins and first-page footer
    # repeat. The letter's own final sectPr becomes the master's new final sectPr.
    lsect = _final_sectpr(lbody)
    for child in list(lbody):
        if child is lsect:
            continue
        mbody.append(copy.deepcopy(child))
    if lsect is not None:
        mbody.append(copy.deepcopy(lsect))


def generate_letters(
    excel_path: str | Path,
    template_path: str | Path,
    map_path: str | Path | None,
    output_path: str | Path,
    preview: bool = False,
) -> int:
    excel_path = Path(excel_path)
    template_path = Path(template_path)
    output_path = Path(output_path)

    case_info, owner_rows = load_case_and_owners(excel_path)
    # An empty minimum-lot-size value (e.g. a workbook saved with a stale
    # formula cache, brief section 20.4) must keep the template's visible
    # X,XXX placeholder rather than silently rendering "of  square feet".
    # A NEW application's proposed minimum is the workbook's 70%-rule
    # calculation (Step 2 K16 -> E4), which the wizard never re-implements;
    # recalculating the workbook (or the wizard's automatic recalc) fills it.
    for _k in ("Proposed_Min_Lot_Size", "Minimum_Lot_Size"):
        if not str(case_info.get(_k) or "").strip():
            case_info[_k] = "X,XXX"
    actual_map_path = resolve_map_path(map_path, excel_path, case_info)
    warnings = validate_inputs(case_info, owner_rows, actual_map_path)
    blocking = [w for w in warnings if w.startswith("Missing Case Info") or w.startswith("No valid owner") or w.startswith("Map file not")]
    if blocking:
        raise ValueError("Cannot generate until these issues are fixed:\n" + "\n".join(blocking))

    if preview:
        owner_rows = [owner_rows[0]]
    map_image_path = convert_pdf_first_page_to_png(actual_map_path)

    # True mail merge. Owner 1 is rendered from the template intact (keeping its
    # native footer relationships). Each subsequent owner is rendered the same
    # way and appended as its OWN section block, so every letter repeats the
    # template's normal body margins AND its first-page council-member footer.
    master_doc = _render_single_owner_letter(template_path, {**case_info, **owner_rows[0]}, map_image_path)
    for owner in owner_rows[1:]:
        letter = _render_single_owner_letter(template_path, {**case_info, **owner}, map_image_path)
        _append_letter_as_own_section(master_doc, letter)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    master_doc.save(str(output_path))
    unlock_docx(output_path)  # draft must be fully editable
    return len(owner_rows)

# ---------------- Wizard v9 additions: Planning Commission Approval + RCA Packet PDF Builder ----------------

def build_pc_context(excel_path: str | Path) -> Dict[str, str]:
    """Build context for the official Planning Commission Approval template."""
    case_info = build_director_context(excel_path)
    if not case_info.get("PC_Approval_Date"):
        case_info["PC_Approval_Date"] = ""
    if not case_info.get("PC_Signer"):
        case_info["PC_Signer"] = "Martha L. Stein, Chair"
    return case_info


def validate_pc_inputs(case_info: Dict[str, str]) -> List[str]:
    warnings = validate_director_inputs(case_info)
    # PC signer/date are optional for document generation; date is often signed later.
    return [w.replace("Director Approval", "Planning Commission Approval") for w in warnings]


def _replace_cell_text_preserve_first_run(cell, replacements: Dict[str, str]) -> None:
    for p in cell.paragraphs:
        text = p.text
        new_text = text
        for old, new in replacements.items():
            if old in new_text:
                new_text = new_text.replace(old, new)
        if new_text != text and p.runs:
            p.runs[0].text = new_text
            p.runs[0].font.color.rgb = RGBColor(0, 0, 0)
            for run in p.runs[1:]:
                run.text = ""


def _patch_pc_template_fields(doc: Document, context: Dict[str, str]) -> None:
    mls = context.get("Minimum_Lot_Size_Number") or lot_size_number(context.get("Minimum_Lot_Size"))
    q = context.get("Qualifying_Lot_Count", "")
    q_words = context.get("Qualifying_Lot_Count_Words", "")
    q_words_cap = context.get("Qualifying_Lot_Count_Words_Cap", q_words[:1].upper() + q_words[1:] if q_words else "")
    total = context.get("Total_Lot_Count", "")
    total_words = context.get("Total_Lot_Count_Words", "")
    q_area = context.get("Qualifying_Area_Percentage", "")
    application_sentence = f"The application is for the {context.get('Director_Block_Description','')}.".replace("..", ".")
    preservation = " ".join([x for x in [context.get("Neighborhood_Age_Statement", ""), context.get("Structure_Age_Statement", ""), context.get("Preservation_Finding", "")] if x]).strip()
    if not preservation:
        preservation = f"The establishment of a {mls} square foot minimum lot size will preserve the lot size character of the area."
    qualification = f"{q_words_cap} ({q}) out of {total_words} ({total}) lots (representing {q_area}% of the application area) are at least {mls} square feet in size."

    replacements = {
        "Special Minimum Lot Size Block No. 319": f"Special Minimum Lot Size Block No. {context.get('Application_Number','')}",
        "The application is for the 1800-1900 block of South Boulevard, north side.": application_sentence,
        "92% of the proposed application area is developed with not more than two SF residential units per property.": f"{context.get('SFR_Percentage','')}% of the proposed application area is developed with not more than two SF residential units per property.",
        "Petition signed by owners of 86% of the proposed SMLSB.": f"Petition signed by owners of {context.get('Support_Percentage','')}% of the proposed SMLSB.",
        "A minimum lot size of 9,450 sq ft exists on six (6) lots in the blockface.": f"A minimum lot size of {mls} square feet exists on {q_words} ({q}) lots in the blockface.",
        "The subdivisions were platted in 1923.  The earliest house originates from the 1930s.  The establishment of a 9,450 square feet minimum lot size will preserve the lot size character of the area.": preservation,
        "Six (6) out of eleven (11) lots (representing 62% of the application area) are at least 9,450 square feet in size.": qualification,
    }

    for p in doc.paragraphs:
        text = p.text
        new_text = text
        for old, new in replacements.items():
            if old in new_text:
                new_text = new_text.replace(old, new)
        if new_text != text and p.runs:
            p.runs[0].text = new_text
            p.runs[0].font.color.rgb = RGBColor(0, 0, 0)
            for run in p.runs[1:]:
                run.text = ""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                _replace_cell_text_preserve_first_run(cell, replacements)


def generate_planning_commission_approval(
    excel_path: str | Path,
    template_path: str | Path,
    output_path: str | Path,
) -> int:
    from core.city_templates import render_updated
    from wizards.mls.templates_2026 import build_context
    if render_updated(template_path, output_path, build_context(excel_path), excel_path, 'mls', 'pc'): return 1
    excel_path = Path(excel_path)
    template_path = Path(template_path)
    output_path = Path(output_path)
    context = build_pc_context(excel_path)
    warnings = validate_pc_inputs(context)
    if warnings:
        raise ValueError("Cannot generate Planning Commission Approval until these issues are fixed:\n" + "\n".join(warnings))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = Document(template_path)
    _patch_pc_template_fields(doc, context)
    doc.save(output_path)
    remove_final_empty_paragraph(output_path)
    unlock_docx(output_path)  # fully editable output
    return 1


def _find_soffice() -> str | None:
    for exe in ["soffice", "libreoffice"]:
        found = shutil.which(exe)
        if found:
            return found
    # Common Windows install locations.
    for p in [
        r"C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        r"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    ]:
        if Path(p).exists():
            return p
    return None


def convert_office_to_pdf(input_path: str | Path, output_dir: str | Path) -> Path:
    """Convert DOCX/XLSX to PDF using LibreOffice when available."""
    input_path = Path(input_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    # LibreOffice when installed, otherwise Microsoft Word/Excel (City PCs).
    from core.office_convert import office_to_pdf
    return office_to_pdf(input_path, output_dir)


def image_to_pdf(input_path: str | Path, output_path: str | Path) -> Path:
    from PIL import Image
    input_path = Path(input_path)
    output_path = Path(output_path)
    img = Image.open(input_path)
    if img.mode in ("RGBA", "LA"):
        bg = Image.new("RGB", img.size, "white")
        bg.paste(img, mask=img.split()[-1])
        img = bg
    else:
        img = img.convert("RGB")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(output_path, "PDF", resolution=150)
    return output_path


def ensure_pdf(input_path: str | Path, work_dir: str | Path) -> Path:
    input_path = Path(input_path)
    work_dir = Path(work_dir)
    suffix = input_path.suffix.lower()
    if suffix == ".pdf":
        return input_path
    if suffix in (".docx", ".doc", ".xlsx", ".xls"):
        return convert_office_to_pdf(input_path, work_dir)
    if suffix in (".png", ".jpg", ".jpeg", ".tif", ".tiff"):
        return image_to_pdf(input_path, work_dir / f"{input_path.stem}.pdf")
    raise ValueError(f"Unsupported packet attachment type: {input_path}")


def _excel_col_letter(n: int) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(n)


STEP2_PROPERTY_HEADER_ROW = 16


def _step2_cell_is_blank(value: Any) -> bool:
    """True when a Step 2 cell holds nothing a planner would consider content.

    Formula-error placeholders and filler zeros count as blank. A property
    description such as "0 (E 25 FT OF LT 14 BLK 5)" is real content, so only an
    exact zero is treated as filler, never a string that merely starts with one.
    """
    if value is None:
        return True
    if isinstance(value, (int, float)):
        return float(value) == 0
    text = str(value).strip()
    return text == "" or text.startswith("#")


def _detect_step2_staff_analysis_range(excel_path: str | Path) -> str:
    """Return the official Staff Analysis export range for Step 2.

    The printable staff-analysis form is columns A:H. Finding the last row is the
    delicate part: the ranked-size columns (E:H) keep running with zero-filler
    rows for a hundred rows past the last parcel, so scanning every column
    produces a huge blank PDF. The parcel table itself is bounded by the
    left-side property columns (description / land use / signed in support),
    which go blank the moment the parcel list ends.

    Earlier versions used row 16 + Total Lots (cell A15) instead. That silently
    dropped parcels, because A15 totals only the SFR, multifamily, commercial,
    and vacant lots, while the table also lists excluded ("EXC") parcels. A15 is
    still honored as a floor so the export can never come up short.
    """
    wb_vals = openpyxl.load_workbook(excel_path, data_only=True)
    ws_vals = wb_vals[STEP2_SHEET]

    last_row = STEP2_PROPERTY_HEADER_ROW
    for row in range(STEP2_PROPERTY_HEADER_ROW + 1, ws_vals.max_row + 1):
        if any(not _step2_cell_is_blank(ws_vals.cell(row, col).value) for col in (1, 2, 3)):
            last_row = row

    total_lots = ws_vals["A15"].value
    try:
        total_lots_int = int(float(total_lots))
    except (TypeError, ValueError):
        total_lots_int = 0
    if total_lots_int > 0:
        last_row = max(last_row, STEP2_PROPERTY_HEADER_ROW + total_lots_int)

    return f"A1:H{last_row}"


def export_step2_staff_analysis_pdf(excel_path: str | Path, output_pdf: str | Path) -> Path:
    """Export Step 2 Staff Analysis to PDF without helper columns or blank formula rows.

    v9 exported the worksheet directly with formulas, which could produce #NAME? errors
    and include helper columns/blank formula rows. v9.2 creates a temporary workbook that
    contains only Step 2's official staff-analysis range (A:H through the last parcel row)
    and replaces formulas with their cached/displayed values before converting to PDF.
    Row heights are then grown so no wrapped description is clipped in the PDF.
    """
    excel_path = Path(excel_path)
    output_pdf = Path(output_pdf)
    tmpdir = Path(tempfile.mkdtemp(prefix="rca_step2_"))
    tmp_xlsx = tmpdir / "Staff_Analysis_Step2.xlsx"

    # Load twice: one copy for styles/layout, one for cached calculated values.
    wb = openpyxl.load_workbook(excel_path, data_only=False)
    wb_vals = openpyxl.load_workbook(excel_path, data_only=True)
    if STEP2_SHEET not in wb.sheetnames or STEP2_SHEET not in wb_vals.sheetnames:
        raise ValueError(f"Could not find required sheet: {STEP2_SHEET}")
    ws = wb[STEP2_SHEET]
    ws_vals = wb_vals[STEP2_SHEET]

    print_area = _detect_step2_staff_analysis_range(excel_path)  # e.g. A1:H38
    from openpyxl.utils.cell import range_boundaries
    min_col, min_row, max_col, max_row = range_boundaries(print_area)

    # Replace formulas in the official printable area with cached/displayed values.
    # This prevents formula errors in the PDF and keeps the form formatting intact.
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row, col)
            val_cell = ws_vals.cell(row, col)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = val_cell.value

    # Clear everything outside the official print area on Step 2, including helper columns.
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            if row < min_row or row > max_row or col < min_col or col > max_col:
                ws.cell(row, col).value = None

    # Remove other sheets so the PDF contains only staff analysis, after formulas are values.
    for sheet in list(wb.sheetnames):
        if sheet != STEP2_SHEET:
            del wb[sheet]
    ws = wb[STEP2_SHEET]

    apply_staff_analysis_page_setup(
        ws, print_area, min_row, max_row, min_col, max_col,
        header_row=STEP2_PROPERTY_HEADER_ROW,
    )

    wb.save(tmp_xlsx)
    pdf = convert_office_to_pdf(tmp_xlsx, tmpdir)
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(pdf, output_pdf)
    return output_pdf


def _get_pdf_writer_class():
    """Return a PDF writer class, installing pypdf if it is missing.

    Some user machines run the wizard before running install_requirements.bat.
    v9.1 attempts a one-time auto-install for pypdf so the RCA Packet PDF button
    does not fail with "No module named 'pypdf'".
    """
    try:
        from pypdf import PdfWriter
        return PdfWriter
    except ModuleNotFoundError:
        import subprocess, sys
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pypdf>=4.0.0"])
            from pypdf import PdfWriter
            return PdfWriter
        except Exception as exc:
            raise ModuleNotFoundError(
                "The RCA Packet PDF builder needs the 'pypdf' package. "
                "Run install_requirements.bat from the wizard folder, then try again. "
                f"Original error: {exc}"
            ) from exc


def merge_pdfs(pdf_paths: List[str | Path], output_pdf: str | Path) -> Path:
    PdfWriter = _get_pdf_writer_class()
    writer = PdfWriter()
    for p in pdf_paths:
        p = Path(p)
        if not p.exists():
            raise FileNotFoundError(p)
        writer.append(str(p))
    output_pdf = Path(output_pdf)
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    with open(output_pdf, "wb") as f:
        writer.write(f)
    return output_pdf


def generate_rca_packet_pdf(
    excel_path: str | Path,
    templates: Dict[str, str | Path],
    map_path: str | Path,
    subdivision_plat_path: str | Path,
    tax_map_path: str | Path,
    hcad_records_path: str | Path | List[str | Path],
    output_pdf: str | Path,
) -> int:
    """Create a combined RCA packet PDF.

    v13 only asks planners for Subdivision Plat, Tax Map, and HCAD Records because
    the wizard can generate everything else. The RCA Checklist has been removed from
    the packet per manager direction. If Approval_Path is Planning Commission Approval,
    the packet uses the PC approval template; otherwise it uses Director Approval.

    v9.3 lets hcad_records_path be a list of PDF paths. Each HCAD PDF is appended
    to the packet in the order selected by the planner. Multi-page HCAD records
    are condensed so each record occupies exactly one page of the packet.
    """
    excel_path = Path(excel_path)
    map_path = Path(map_path)
    output_pdf = Path(output_pdf)
    case_info, _owners = load_case_and_owners(excel_path)
    approval_path = case_info.get("Approval_Path", "Director Approval")
    from core.packet_work import packet_workspace
    with packet_workspace() as work_dir:

        routing_docx = work_dir / "01_MLS_Routing_Form.docx"
        generate_routing_form(excel_path, templates["routing"], routing_docx)

        pdfs: List[Path] = [
            routing_pdf_one_page(ensure_pdf(routing_docx, work_dir), work_dir),
        ]
        if approval_path == "Planning Commission Approval":
            pc_docx = work_dir / "03_Planning_Commission_Approval.docx"
            generate_planning_commission_approval(excel_path, templates["pc"], pc_docx)
            pdfs.append(ensure_pdf(pc_docx, work_dir))
        else:
            director_docx = work_dir / "03_Director_Approval.docx"
            generate_director_approval(excel_path, templates["director"], director_docx)
            pdfs.append(ensure_pdf(director_docx, work_dir))

        pdfs.append(ensure_pdf(map_path, work_dir))
        staff_pdf = work_dir / "05_Staff_Analysis_Step2.pdf"
        export_step2_staff_analysis_pdf(excel_path, staff_pdf)
        pdfs.append(staff_pdf)
        if str(subdivision_plat_path or "").strip():
            pdfs.append(ensure_pdf(subdivision_plat_path, work_dir))
        pdfs.append(ensure_pdf(tax_map_path, work_dir))

        if isinstance(hcad_records_path, (list, tuple)):
            hcad_paths = [Path(p) for p in hcad_records_path if str(p).strip()]
        else:
            # Backward compatible: allow a single path or a delimiter-separated field from the UI.
            raw = str(hcad_records_path or "")
            hcad_paths = [Path(p) for p in raw.split("||") if p.strip()]
        pdfs.extend(condense_hcad_pdfs([ensure_pdf(p, work_dir) for p in hcad_paths], work_dir))

        merge_pdfs(pdfs, output_pdf)
        return len(pdfs)

# --- v10 Draft Ordinance Generator -------------------------------------------------
REQUIRED_ORDINANCE_FIELDS = [
    "Application_Number", "Minimum_Lot_Size_Number", "Total_Lot_Count",
    "Support_Percentage", "Director_Name"
]


def _strip_null_text(value: Any) -> str:
    s = normalize_value(value)
    return "" if s in ("<Null>", "NULL", "None") else s


def _combine_legal_description(row: Dict[str, str]) -> str:
    parts = [row.get(f"Legal_Dscr_{i}", "") for i in range(1, 5)]
    parts = [_strip_null_text(p) for p in parts]
    return " ".join([p for p in parts if p]).strip()


def _parse_lot_block_subdivision(legal: str) -> Tuple[str, str, str]:
    # Block = first BLK token if present.
    block = ""
    m = re.search(r"\bBLK\s+([A-Za-z0-9\-]+)", legal, flags=re.I)
    if m:
        block = m.group(1)
    # Lot/tract = legal text up to BLK if present; otherwise full legal.
    lot = legal
    if m:
        lot = legal[:m.start()].strip()
    lot = re.sub(r"\s+", " ", lot).strip()
    # Subdivision = anything after the BLK number if present, otherwise blank.
    subdivision = ""
    if m:
        subdivision = legal[m.end():].strip(" ,;-")
    # If legal desc is multiline where subdivision is in later legal field, this catches it.
    subdivision = re.sub(r"\s+", " ", subdivision).strip()
    return lot, block, subdivision


def _read_property_rows_for_ordinance(excel_path: str | Path) -> List[Dict[str, str]]:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if STEP1_SHEET not in wb.sheetnames:
        raise ValueError(f"Workbook must contain the '{STEP1_SHEET}' tab.")
    ws = wb[STEP1_SHEET]
    headers = [normalize_value(c.value) for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(headers) if h}
    def cell(r, name):
        return _strip_null_text(ws.cell(r, idx[name]).value) if name in idx else ""
    rows = []
    for r in range(2, ws.max_row + 1):
        address = cell(r, "Site_addr_1") or cell(r, "Property Description")
        hcad = cell(r, "TAX_ID")
        if not address and not hcad:
            continue
        legal_parts = {f"Legal_Dscr_{i}": cell(r, f"Legal_Dscr_{i}") for i in range(1, 5)}
        legal = _combine_legal_description(legal_parts)
        lot, block, subdiv = _parse_lot_block_subdivision(legal)
        if not subdiv:
            # Prefer the last legal line that does not look like a lot/block continuation.
            candidates = [legal_parts[f"Legal_Dscr_{i}"] for i in range(4, 0, -1)]
            for c in candidates:
                if c and not re.search(r"\b(LT|LOT|TR|TRS|BLK|BLOCK)\b", c, flags=re.I):
                    subdiv = c
                    break
        rows.append({
            "ADDRESS": address,
            "HCAD": hcad,
            "LEGAL_DESCRIPTION": legal,
            "LOT_TRACT": lot,
            "BLOCK": block,
            "SUBDIVISION": subdiv,
        })
    return rows


def _ordinance_side_block_phrase(sides: str, block: str) -> str:
    """Build the ordinance-only side/block phrase without changing Case Info values.

    Case Info stores the clean reusable side phrase, for example
    "north and south sides". The ordinance sentence needs the grammar
    "north and south sides of the 4800-4900 block..." while routing and
    approval forms need the clean phrase. Keep that wording here/template-side
    instead of requiring Case Info B46 to contain "of the".
    """
    sides = normalize_value(sides)
    block = normalize_value(block)
    if sides and block:
        return f"{sides} of the {block}"
    return sides or block


def _add_ordinance_template_fields(context: Dict[str, str]) -> None:
    """Add component fields used by the ordinance template.

    These fields allow the ordinance template to supply ordinance-specific
    grammar while routing, Director Approval, and PC Approval continue using
    the clean Side_Of_Street value.
    """
    side_block = _ordinance_side_block_phrase(context.get("Side_Of_Street", ""), context.get("Block_Description", ""))
    cross = context.get("Cross_Streets", "")
    context["Ordinance_Side_Block_Lower"] = side_block
    context["Ordinance_Side_Block_Upper"] = side_block.upper()
    context["Cross_Streets_Upper"] = cross.upper()


def _ordinance_area_from_case(case_info: Dict[str, str]) -> Tuple[str, str]:
    block = case_info.get("Block_Description", "")
    sides = case_info.get("Side_Of_Street", "")
    cross = case_info.get("Cross_Streets", "")
    lower = _ordinance_side_block_phrase(sides, block)
    if cross:
        lower = f"{lower}, {cross}".strip(", ")
    upper = lower.upper()
    return upper, lower


def build_ordinance_context(excel_path: str | Path) -> Dict[str, str]:
    context = build_director_context(excel_path)
    area_upper, area_lower = _ordinance_area_from_case(context)
    context["Ordinance_Area_Upper"] = area_upper
    context["Ordinance_Area_Lower"] = area_lower
    _add_ordinance_template_fields(context)
    context.setdefault("Ordinance_Year", "2026")
    context.setdefault("Legal_File_No", "")
    context.setdefault("Director_Name", "Vonn Tran")
    # Some worksheets return strings like "6380-Square Foot". Use just the number for ordinance text.
    n = lot_size_number(context.get("Minimum_Lot_Size_Number") or context.get("Minimum_Lot_Size") or context.get("Proposed_Min_Lot_Size"))
    context["Minimum_Lot_Size_Number"] = n
    try:
        context["Minimum_Lot_Size_Words"] = number_word(int(n.replace(",", "")))
    except Exception:
        context["Minimum_Lot_Size_Words"] = n
    return context


def validate_ordinance_inputs(context: Dict[str, str]) -> List[str]:
    warnings = []
    for field in REQUIRED_ORDINANCE_FIELDS:
        if not context.get(field):
            warnings.append(f"Missing field for Draft Ordinance: {field}")
    if not context.get("Ordinance_Area_Lower"):
        warnings.append("Missing ordinance application area description.")
    return warnings


def _insert_paragraph_after(paragraph: Paragraph, text: str = "") -> Paragraph:
    from docx.oxml import OxmlElement
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    return new_para


def _find_paragraph(doc: Document, exact_text: str) -> Paragraph | None:
    for p in doc.paragraphs:
        if p.text.strip() == exact_text:
            return p
    return None


def _insert_map_after_header(doc: Document, map_path: Path) -> None:
    header = _find_paragraph(doc, "BOUNDARY MAP")
    if not header:
        return
    image = convert_pdf_first_page_to_png(map_path)
    p = _insert_paragraph_after(header, "")
    run = p.add_run()
    run.add_picture(str(image), width=Inches(6.25))
    p.alignment = 1


def _insert_property_descriptions(doc: Document, rows: List[Dict[str, str]]) -> None:
    header = _find_paragraph(doc, "PROPERTY DESCRIPTIONS")
    if not header:
        return
    cur = header
    for row in rows:
        legal = row.get("LEGAL_DESCRIPTION", "")
        if legal:
            cur = _insert_paragraph_after(cur, legal)


def _insert_property_list_table(doc: Document, rows: List[Dict[str, str]]) -> None:
    header = _find_paragraph(doc, "PROPERTY LIST")
    if not header:
        return
    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    for cell, text in zip(hdr, ["ADDRESS", "HCAD", "LOT/TRACT", "BLOCK", "SUBDIVISION"]):
        cell.text = text
        for p in cell.paragraphs:
            for r in p.runs:
                r.bold = True
    for row in rows:
        cells = table.add_row().cells
        cells[0].text = row.get("ADDRESS", "")
        cells[1].text = row.get("HCAD", "")
        cells[2].text = row.get("LOT_TRACT", "")
        cells[3].text = row.get("BLOCK", "")
        cells[4].text = row.get("SUBDIVISION", "")
    # Move the table after the PROPERTY LIST paragraph.
    tbl = table._tbl
    doc.element.body.remove(tbl)
    header._p.addnext(tbl)


def generate_draft_ordinance(
    excel_path: str | Path,
    template_path: str | Path,
    map_path: str | Path,
    output_path: str | Path,
) -> int:
    excel_path = Path(excel_path)
    template_path = Path(template_path)
    output_path = Path(output_path)
    map_path = Path(map_path)
    context = build_ordinance_context(excel_path)
    warnings = validate_ordinance_inputs(context)
    if warnings:
        raise ValueError("Cannot generate Draft Ordinance until these issues are fixed:\n" + "\n".join(warnings))
    doc = Document(template_path)
    fill_elements(doc, None, context, Path("__no_map_needed__.png"))
    rows = _read_property_rows_for_ordinance(excel_path)
    _insert_property_descriptions(doc, rows)
    if map_path.exists():
        _insert_map_after_header(doc, map_path)
    _insert_property_list_table(doc, rows)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    unlock_docx(output_path)  # fully editable output
    return 1

# Patch number-word conversion for ordinance lot sizes beyond 50.
def number_to_words(n: int) -> str:
    ones = ["zero","one","two","three","four","five","six","seven","eight","nine","ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen"]
    tens = ["","","twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety"]
    def under1000(x):
        parts=[]
        if x>=100:
            parts.append(ones[x//100]+" hundred")
            x%=100
        if x>=20:
            parts.append(tens[x//10] + (("-"+ones[x%10]) if x%10 else ""))
        elif x>0 or not parts:
            parts.append(ones[x])
        return " ".join(parts)
    if n < 1000:
        return under1000(n)
    if n < 1000000:
        rem = n % 1000
        base = under1000(n//1000) + " thousand"
        return base + ((" " + under1000(rem)) if rem else "")
    return str(n)

# Override build_ordinance_context with better number words.
_old_build_ordinance_context = build_ordinance_context
def build_ordinance_context(excel_path: str | Path) -> Dict[str, str]:
    context = _old_build_ordinance_context(excel_path)
    n = lot_size_number(context.get("Minimum_Lot_Size_Number") or context.get("Minimum_Lot_Size") or context.get("Proposed_Min_Lot_Size"))
    context["Minimum_Lot_Size_Number"] = n
    try:
        context["Minimum_Lot_Size_Words"] = number_to_words(int(n.replace(",", "")))
    except Exception:
        context["Minimum_Lot_Size_Words"] = n
    return context

# v10 robustness patch: if Step 2 formulas are not cached by Excel, compute ordinance basics from Step 1.
def _read_property_rows_with_lot_size(excel_path: str | Path) -> List[Dict[str, str]]:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[STEP1_SHEET]
    headers = [normalize_value(c.value) for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(headers) if h}
    rows=[]
    for r in range(2, ws.max_row+1):
        addr = _strip_null_text(ws.cell(r, idx.get("Site_addr_1", 1)).value) if "Site_addr_1" in idx else ""
        if not addr: continue
        lot_size = _strip_null_text(ws.cell(r, idx.get("Lot Size", 1)).value) if "Lot Size" in idx else ""
        rows.append({"address": addr, "lot_size": lot_size})
    return rows

_prev_build_ordinance_context = build_ordinance_context
def build_ordinance_context(excel_path: str | Path) -> Dict[str, str]:
    context = _prev_build_ordinance_context(excel_path)
    rows = _read_property_rows_with_lot_size(excel_path)
    if not context.get("Total_Lot_Count"):
        context["Total_Lot_Count"] = str(len(rows))
        context["Total_Lot_Count_Words"] = number_word(len(rows)) if len(rows) <= 50 else number_to_words(len(rows))
    if not context.get("Support_Percentage"):
        context["Support_Percentage"] = "XX"
    min_num = lot_size_number(context.get("Minimum_Lot_Size_Number") or context.get("Minimum_Lot_Size") or context.get("Proposed_Min_Lot_Size"))
    context["Minimum_Lot_Size_Number"] = min_num
    try:
        min_i = int(min_num.replace(",", ""))
        context["Minimum_Lot_Size_Words"] = number_to_words(min_i)
        if not context.get("Qualifying_Lot_Count"):
            q = 0
            for row in rows:
                try:
                    if int(str(row.get("lot_size", "")).replace(",", "")) >= min_i:
                        q += 1
                except Exception:
                    pass
            context["Qualifying_Lot_Count"] = str(q)
            context["Qualifying_Lot_Count_Words"] = number_word(q) if q <= 50 else number_to_words(q)
    except Exception:
        pass
    return context

# Override validation: allow Support_Percentage to be a planner fill-in placeholder if formulas are not cached.
def validate_ordinance_inputs(context: Dict[str, str]) -> List[str]:
    warnings = []
    for field in ["Application_Number", "Minimum_Lot_Size_Number", "Total_Lot_Count", "Director_Name"]:
        if not context.get(field):
            warnings.append(f"Missing field for Draft Ordinance: {field}")
    if not context.get("Ordinance_Area_Lower"):
        warnings.append("Missing ordinance application area description.")
    return warnings

# v10 patch: Step 1 often stores lot size in Total_Land_Area rather than the derived Lot Size column.
def _read_property_rows_with_lot_size(excel_path: str | Path) -> List[Dict[str, str]]:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[STEP1_SHEET]
    headers = [normalize_value(c.value) for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(headers) if h}
    rows=[]
    for r in range(2, ws.max_row+1):
        addr = _strip_null_text(ws.cell(r, idx.get("Site_addr_1", 1)).value) if "Site_addr_1" in idx else ""
        if not addr: continue
        lot_size = ""
        for name in ["Lot Size", "Total_Land_Area"]:
            if name in idx:
                lot_size = _strip_null_text(ws.cell(r, idx[name]).value)
                if lot_size:
                    break
        rows.append({"address": addr, "lot_size": lot_size})
    return rows

_prev2_build_ordinance_context = build_ordinance_context
def build_ordinance_context(excel_path: str | Path) -> Dict[str, str]:
    context = _prev2_build_ordinance_context(excel_path)
    rows = _read_property_rows_with_lot_size(excel_path)
    min_num = lot_size_number(context.get("Minimum_Lot_Size_Number") or context.get("Minimum_Lot_Size") or context.get("Proposed_Min_Lot_Size"))
    context["Minimum_Lot_Size_Number"] = min_num
    try:
        min_i = int(min_num.replace(",", ""))
        q = 0
        for row in rows:
            try:
                if int(str(row.get("lot_size", "")).replace(",", "")) >= min_i:
                    q += 1
            except Exception:
                pass
        # Prefer computed value from Step 1 when the formula-derived value is blank or zero.
        if not context.get("Qualifying_Lot_Count") or context.get("Qualifying_Lot_Count") == "0":
            context["Qualifying_Lot_Count"] = str(q)
            context["Qualifying_Lot_Count_Words"] = number_word(q) if q <= 50 else number_to_words(q)
    except Exception:
        pass
    return context

def trim_blank_tail_after_last_table(docx_path: str | Path) -> None:
    import zipfile, tempfile, shutil
    from lxml import etree
    docx_path = Path(docx_path)
    tmp = Path(tempfile.mkdtemp(prefix="trim_docx_"))
    with zipfile.ZipFile(docx_path) as z:
        z.extractall(tmp)
    xmlp = tmp / "word/document.xml"
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    tree = etree.parse(str(xmlp))
    body = tree.getroot().find("w:body", namespaces=ns)
    if body is not None:
        children = list(body)
        # find the last table in the body
        last_tbl_idx = None
        for i, child in enumerate(children):
            if etree.QName(child).localname == "tbl":
                last_tbl_idx = i
        if last_tbl_idx is not None:
            # remove blank paragraphs after the last table, before sectPr
            for child in list(body)[last_tbl_idx+1:]:
                lname = etree.QName(child).localname
                if lname == "sectPr":
                    continue
                texts = child.xpath('.//w:t/text()', namespaces=ns)
                drawings = child.xpath('.//w:drawing', namespaces=ns)
                if lname == "p" and not ''.join(texts).strip() and not drawings:
                    body.remove(child)
        tree.write(str(xmlp), xml_declaration=True, encoding="UTF-8", standalone="yes")
        with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as z:
            for file in tmp.rglob("*"):
                if file.is_file():
                    z.write(file, file.relative_to(tmp).as_posix())
    shutil.rmtree(tmp, ignore_errors=True)

# Override generator to trim trailing blank page after Exhibit C table.
_prev_generate_draft_ordinance = generate_draft_ordinance
def generate_draft_ordinance(excel_path: str | Path, template_path: str | Path, map_path: str | Path, output_path: str | Path) -> int:
    result = _prev_generate_draft_ordinance(excel_path, template_path, map_path, output_path)
    trim_blank_tail_after_last_table(output_path)
    return result

# v10 validation convenience: sample workbooks may not have Excel-cached formula results.
_prev_build_director_context = build_director_context
def build_director_context(excel_path: str | Path) -> Dict[str, str]:
    context = _prev_build_director_context(excel_path)
    ord_ctx = build_ordinance_context(excel_path)
    for k in ["Support_Percentage", "Minimum_Lot_Size_Number", "Qualifying_Lot_Count", "Qualifying_Lot_Count_Words", "Qualifying_Area_Percentage", "Total_Lot_Count", "Total_Lot_Count_Words"]:
        if not context.get(k) and ord_ctx.get(k):
            context[k] = ord_ctx[k]
    if not context.get("SFR_Percentage"):
        context["SFR_Percentage"] = "XX"
    return context

# --- v10.1 calculation patch -------------------------------------------------------
# Step 2 already contains the signed-support rows and calculation logic. When Excel
# has not cached formula results, openpyxl cannot read formulas as values. These
# helpers compute the same case facts from Step 1 + Step 2 visible planner inputs.
def _case_valid_step1_rows(excel_path: str | Path) -> list[dict[str, Any]]:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if STEP1_SHEET not in wb.sheetnames:
        return []
    ws = wb[STEP1_SHEET]
    headers = [normalize_value(c.value) for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(headers) if h}
    rows: list[dict[str, Any]] = []
    for r in range(2, ws.max_row + 1):
        addr = _strip_null_text(ws.cell(r, idx.get("Site_addr_1", 1)).value) if "Site_addr_1" in idx else ""
        tax = _strip_null_text(ws.cell(r, idx.get("TAX_ID", 1)).value) if "TAX_ID" in idx else ""
        raw_area = ws.cell(r, idx.get("Total_Land_Area", 1)).value if "Total_Land_Area" in idx else None
        try:
            area = float(str(raw_area).replace(",", ""))
        except Exception:
            area = 0.0
        if not addr and not tax:
            continue
        if area <= 0:
            continue
        land = _strip_null_text(ws.cell(r, idx.get("group_dscr", 1)).value) if "group_dscr" in idx else ""
        rows.append({"excel_row": r, "address": addr, "tax_id": tax, "area": area, "land_use": land})
    return rows


def _compute_step2_support_from_visible_y(excel_path: str | Path) -> dict[str, str]:
    rows = _case_valid_step1_rows(excel_path)
    total_area = sum(r["area"] for r in rows)
    signed_area = 0.0
    if total_area <= 0:
        return {}
    wb_formula = openpyxl.load_workbook(excel_path, data_only=False)
    if STEP2_SHEET not in wb_formula.sheetnames:
        return {}
    ws2 = wb_formula[STEP2_SHEET]
    for i, row in enumerate(rows, start=17):
        signed = normalize_value(ws2.cell(i, 3).value).strip().upper()  # Step 2 col C, Signed in Support
        if signed in {"Y", "YES", "TRUE", "1"}:
            signed_area += row["area"]
    return {
        "Support_Percentage": str(int(round((signed_area / total_area) * 100))),
        "Support_Area": str(int(round(signed_area))),
        "Total_Area": str(int(round(total_area))),
    }


def _compute_step2_min_lot_values_from_step1(excel_path: str | Path) -> dict[str, str]:
    rows = _case_valid_step1_rows(excel_path)
    if not rows:
        return {}
    total_area = sum(r["area"] for r in rows)
    sorted_areas = sorted([r["area"] for r in rows], reverse=True)
    running = 0.0
    mls = sorted_areas[-1]
    q_count = len(sorted_areas)
    q_pct = 1.0
    for n, area in enumerate(sorted_areas, start=1):
        running += area
        if total_area and running / total_area >= 0.70:
            mls = area
            q_count = n
            q_pct = running / total_area
            break
    total_lots = len(rows)
    sfr_lots = sum(1 for r in rows if "single" in str(r.get("land_use", "")).lower())
    mls_i = int(round(mls))
    # Note: Qualifying_Lot_Count is NOT set here. The final Step 2 reader counts each eligible
    # property meeting the minimum, or uses Lots_Meeting_Minimum when entered.
    # The 70%-rank q_count above is only the
    # minimum-lot-size determination and is a different number, so emitting it
    # as the qualifying count would put the wrong value in the documents. The
    # 70%-area percentage (q_pct) is a genuine calculated value and is kept.
    return {
        "Total_Lot_Count": str(total_lots),
        "Total_Lot_Count_Words": number_word(total_lots) if total_lots <= 50 else number_to_words(total_lots),
        "Minimum_Lot_Size_Number": f"{mls_i:,}",
        "Minimum_Lot_Size_Words": number_to_words(mls_i),
        "Qualifying_Area_Percentage": str(int(round(q_pct * 100))),
        "SFR_Percentage": str(int(round((sfr_lots / total_lots) * 100))) if total_lots else "",
    }


def _build_calculation_context_v101(excel_path: str | Path) -> dict[str, str]:
    case_info, _owners = load_case_and_owners(excel_path)
    # Prefer cached Step 2 values where available, then fill gaps with calculated values.
    try:
        cached = read_director_defaults_from_step2(excel_path)
        for k, v in cached.items():
            if v not in (None, ""):
                case_info[k] = normalize_value(v)
    except Exception:
        pass
    # Manual calculations should override blank/placeholder values produced when formulas are not cached.
    manual_min = _compute_step2_min_lot_values_from_step1(excel_path)
    for k, v in manual_min.items():
        if v and (not case_info.get(k) or case_info.get(k) in {"0", "0-Square Foot", "XX", "#NAME?"}):
            case_info[k] = v
    manual_support = _compute_step2_support_from_visible_y(excel_path)
    if manual_support.get("Support_Percentage"):
        case_info["Support_Percentage"] = manual_support["Support_Percentage"]
        case_info["Support_Area"] = manual_support.get("Support_Area", "")
        case_info["Total_Area"] = manual_support.get("Total_Area", "")
    # Normalize minimum lot size value.
    n = lot_size_number(case_info.get("Minimum_Lot_Size_Number") or case_info.get("Minimum_Lot_Size") or case_info.get("Proposed_Min_Lot_Size"))
    if n and n not in {"0"}:
        case_info["Minimum_Lot_Size_Number"] = n
        # The owner letter merges Proposed_Min_Lot_Size / Minimum_Lot_Size directly,
        # and a manual fallback above can set these to a "NNNN-Square Foot" label.
        # Keep them as the clean number so the letter never renders "-Square Foot".
        case_info["Proposed_Min_Lot_Size"] = n
        case_info["Minimum_Lot_Size"] = n
        try:
            case_info["Minimum_Lot_Size_Words"] = number_to_words(int(n.replace(",", "")))
        except Exception:
            pass
    director_block = ", ".join([x for x in [case_info.get("Block_Description", ""), case_info.get("Side_Of_Street", "")] if x])
    cross = case_info.get("Cross_Streets", "")
    if cross:
        director_block = f"{director_block} {cross}".strip()
    case_info["Director_Block_Description"] = director_block
    case_info.setdefault("Director_Name", "Vonn Tran")
    case_info.setdefault("Director_Approval_Date", "")
    if not case_info.get("Preservation_Finding"):
        mls = case_info.get("Minimum_Lot_Size_Number", "")
        case_info["Preservation_Finding"] = f"The establishment of a {mls} square foot minimum lot size will preserve the lot size character of the area." if mls else ""
    return case_info


def build_director_context(excel_path: str | Path) -> Dict[str, str]:
    return _build_calculation_context_v101(excel_path)


def build_ordinance_context(excel_path: str | Path) -> Dict[str, str]:
    context = _build_calculation_context_v101(excel_path)
    area_upper, area_lower = _ordinance_area_from_case(context)
    context["Ordinance_Area_Upper"] = area_upper.upper()
    context["Ordinance_Area_Lower"] = area_lower
    _add_ordinance_template_fields(context)
    context.setdefault("Ordinance_Year", "2026")
    context.setdefault("Legal_File_No", "")
    context.setdefault("Director_Name", "Vonn Tran")
    return context

# --- v10.1.1 ordinal/text sync patch ---------------------------------------------
# Keep numeric and word fields synchronized when calculated values override stale cached formula text.
_prev_v101_calc_context = _build_calculation_context_v101

def _build_calculation_context_v101(excel_path: str | Path) -> dict[str, str]:
    case_info = _prev_v101_calc_context(excel_path)
    manual_min = _compute_step2_min_lot_values_from_step1(excel_path)
    # Always synchronize these display fields as a set so we never get "zero (11)".
    for k in [
        "Total_Lot_Count", "Total_Lot_Count_Words", "Minimum_Lot_Size_Number", "Minimum_Lot_Size_Words",
        "Qualifying_Lot_Count", "Qualifying_Lot_Count_Words", "Qualifying_Lot_Count_Words_Cap",
        "Qualifying_Area_Percentage", "SFR_Percentage",
    ]:
        if manual_min.get(k):
            case_info[k] = manual_min[k]
    manual_support = _compute_step2_support_from_visible_y(excel_path)
    if manual_support.get("Support_Percentage"):
        case_info["Support_Percentage"] = manual_support["Support_Percentage"]
        case_info["Support_Area"] = manual_support.get("Support_Area", "")
        case_info["Total_Area"] = manual_support.get("Total_Area", "")
    return case_info

# Rebind public context functions to the synchronized implementation.
def build_director_context(excel_path: str | Path) -> Dict[str, str]:
    return _build_calculation_context_v101(excel_path)

def build_ordinance_context(excel_path: str | Path) -> Dict[str, str]:
    context = _build_calculation_context_v101(excel_path)
    area_upper, area_lower = _ordinance_area_from_case(context)
    context["Ordinance_Area_Upper"] = area_upper.upper()
    context["Ordinance_Area_Lower"] = area_lower
    _add_ordinance_template_fields(context)
    context.setdefault("Ordinance_Year", "2026")
    context.setdefault("Legal_File_No", "")
    context.setdefault("Director_Name", "Vonn Tran")
    return context


# --- v13.2 support percentage source-of-truth patch ------------------------------
# The support percentage shown in the Director Approval, Planning Commission
# Approval, and Draft Ordinance must come from Step 2 of the calculation sheet.
# Earlier fallback logic could recompute support from Step 1 and Step 2 signs if
# Excel formula values were not cached. That is useful as a fallback, but it must
# not override the planner-visible Step 2 support percentage when Step 2 has it.
def _normalize_percent_text_for_template(value: Any) -> str:
    raw = normalize_value(value).strip()
    if not raw or raw in {"#NAME?", "#VALUE!", "#REF!", "XX"}:
        return ""
    raw = raw.replace("%", "").strip()
    try:
        num = float(raw.replace(",", ""))
        if abs(num) <= 1:
            num *= 100
        return str(int(round(num)))
    except Exception:
        m = re.search(r"(-?\d+(?:\.\d+)?)\s*%?", raw)
        if not m:
            return ""
        try:
            num = float(m.group(1))
            if abs(num) <= 1:
                num *= 100
            return str(int(round(num)))
        except Exception:
            return m.group(1)


def _read_step2_visible_support_percentage(excel_path: str | Path) -> str:
    """Read the visible Step 2 support percentage from the staff-analysis form.

    The official Step 2 form has an Evidence of Support section ending with a
    highlighted percent cell. This function first checks the known area, then
    scans the Evidence of Support block for the best percent candidate. It only
    returns a value if it appears to be a real percent, not helper text.
    """
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception:
        return ""
    if STEP2_SHEET not in wb.sheetnames:
        return ""
    ws = wb[STEP2_SHEET]

    # First try likely visible percent cells in the top support block.
    for addr in ("H8", "G8", "F8", "H7", "G7", "F7"):
        val = _normalize_percent_text_for_template(ws[addr].value)
        if val and 0 <= float(val) <= 100:
            return val

    # Locate the Evidence of Support block and scan nearby cells for percent values.
    start_row = None
    for row in range(1, min(ws.max_row, 40) + 1):
        row_text = " ".join(normalize_value(ws.cell(row, col).value) for col in range(1, min(ws.max_column, 12) + 1))
        if "Evidence of Support" in row_text:
            start_row = row
            break
    if start_row is None:
        start_row = 1

    candidates: list[tuple[int, int, str]] = []
    for row in range(start_row, min(start_row + 10, ws.max_row) + 1):
        for col in range(1, min(ws.max_column, 12) + 1):
            cell = ws.cell(row, col)
            val = _normalize_percent_text_for_template(cell.value)
            if not val:
                continue
            try:
                f = float(val)
            except Exception:
                continue
            if 0 <= f <= 100:
                # Prefer cells farther right and lower in the support block, where the highlighted percent normally sits.
                candidates.append((row, col, val))
    if not candidates:
        return ""
    # Prefer the right-most/lower-most candidate in the Evidence of Support block, which avoids total-area values.
    candidates.sort(key=lambda x: (x[1], x[0]), reverse=True)
    return candidates[0][2]


_prev_v132_calc_context = _build_calculation_context_v101

def _build_calculation_context_v101(excel_path: str | Path) -> dict[str, str]:
    case_info = _prev_v132_calc_context(excel_path)
    visible_support = _read_step2_visible_support_percentage(excel_path)
    if visible_support:
        case_info["Support_Percentage"] = visible_support
    return case_info

# Rebind public context functions again so all modules use Step 2 visible support.
def build_director_context(excel_path: str | Path) -> Dict[str, str]:
    return _build_calculation_context_v101(excel_path)

def build_pc_context(excel_path: str | Path) -> Dict[str, str]:
    case_info = build_director_context(excel_path)
    if not case_info.get("PC_Approval_Date"):
        case_info["PC_Approval_Date"] = ""
    if not case_info.get("PC_Signer"):
        case_info["PC_Signer"] = "Martha L. Stein, Chair"
    return case_info

def build_ordinance_context(excel_path: str | Path) -> Dict[str, str]:
    context = _build_calculation_context_v101(excel_path)
    area_upper, area_lower = _ordinance_area_from_case(context)
    context["Ordinance_Area_Upper"] = area_upper.upper()
    context["Ordinance_Area_Lower"] = area_lower
    _add_ordinance_template_fields(context)
    context.setdefault("Ordinance_Year", "2026")
    context.setdefault("Legal_File_No", "")
    context.setdefault("Director_Name", "Vonn Tran")
    return context


# --- v13.3 ordinance font normalization patch -----------------------------------
def _normalize_draft_ordinance_fonts(doc: Document) -> None:
    """Keep generated ordinance text visually consistent with the City template.

    Some placeholder runs in the ordinance template can inherit Word's Normal style
    after replacement, which makes inserted facts render smaller/different from the
    surrounding Arial 12 ordinance text. This patch only fills in missing run-level
    font metadata; it does not change bolding, spacing, or any run that already has
    an explicit font/size.
    """
    for paragraph in iter_wrapped_paragraphs_and_tables(doc):
        for run in paragraph.runs:
            if run.text:
                if run.font.name is None:
                    run.font.name = "Arial"
                if run.font.size is None:
                    run.font.size = Pt(12)

_prev_generate_draft_ordinance_v133 = generate_draft_ordinance
def generate_draft_ordinance(excel_path: str | Path, template_path: str | Path, map_path: str | Path, output_path: str | Path) -> int:
    # Re-run the ordinance generation with the same template/content logic, then
    # normalize only missing font metadata before final save.
    excel_path = Path(excel_path)
    template_path = Path(template_path)
    output_path = Path(output_path)
    map_path = Path(map_path)
    context = build_ordinance_context(excel_path)
    warnings = validate_ordinance_inputs(context)
    if warnings:
        raise ValueError("Cannot generate Draft Ordinance until these issues are fixed:\n" + "\n".join(warnings))
    doc = Document(template_path)
    fill_elements(doc, None, context, Path("__no_map_needed__.png"))
    rows = _read_property_rows_for_ordinance(excel_path)
    _insert_property_descriptions(doc, rows)
    if map_path.exists():
        _insert_map_after_header(doc, map_path)
    _insert_property_list_table(doc, rows)
    _normalize_draft_ordinance_fonts(doc)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    trim_blank_tail_after_last_table(output_path)
    unlock_docx(output_path)  # draft must be fully editable
    return 1


# --- v13.6 SFR percentage source-of-truth patch ---------------------------------
# Director/PC approvals must use the visible Step 2 SFR percentage, not the
# fallback Step 1 property-count calculation. In MLS 905, Step 2 shows 14 / 22 =
# 64%, while the fallback counted 14 / 21 = 67%. Step 2 is the source of truth.
def _read_step2_visible_sfr_percentage(excel_path: str | Path) -> str:
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if STEP2_SHEET not in wb.sheetnames:
            return ""
        ws = wb[STEP2_SHEET]
        # Standard MLS Step 2 SFR percentage cell.
        val = ws["H11"].value
        pct_val = pct(val)
        if pct_val:
            return pct_val
    except Exception:
        return ""
    return ""

_prev_v136_calc_context = _build_calculation_context_v101

def _build_calculation_context_v101(excel_path: str | Path) -> dict[str, str]:
    case_info = _prev_v136_calc_context(excel_path)
    visible_sfr = _read_step2_visible_sfr_percentage(excel_path)
    if visible_sfr:
        case_info["SFR_Percentage"] = visible_sfr
    return case_info

# Rebind public context functions so Director Approval, PC Approval, Routing Form,
# Ordinance, and packet generation all see the Step 2 SFR percentage.
def build_director_context(excel_path: str | Path) -> Dict[str, str]:
    return _build_calculation_context_v101(excel_path)

def build_pc_context(excel_path: str | Path) -> Dict[str, str]:
    case_info = build_director_context(excel_path)
    if not case_info.get("PC_Approval_Date"):
        case_info["PC_Approval_Date"] = ""
    if not case_info.get("PC_Signer"):
        case_info["PC_Signer"] = "Martha L. Stein, Chair"
    return case_info

def build_ordinance_context(excel_path: str | Path) -> Dict[str, str]:
    context = _build_calculation_context_v101(excel_path)
    area_upper, area_lower = _ordinance_area_from_case(context)
    context["Ordinance_Area_Upper"] = area_upper.upper()
    context["Ordinance_Area_Lower"] = area_lower
    _add_ordinance_template_fields(context)
    context.setdefault("Ordinance_Year", "2026")
    context.setdefault("Legal_File_No", "")
    context.setdefault("Director_Name", "Vonn Tran")
    return context

# --- v13.7 Step 2/Case Info source-of-truth audit patch -------------------------
# Calculation facts used by Routing Form, Director Approval, PC Approval, Draft
# Ordinance, and RCA packet generation must come from the official Step 2 staff
# analysis or the Case Info review tab. Earlier fallback code could compute total
# lot count from Step 1 positive-area rows; that caused MLS 905 to show 21 lots
# when Step 2 excluded two lots and showed 20 official lots. This patch re-applies
# the Step 2/Case Info source of truth after any fallback calculations.

def _clean_number_text(value: Any) -> str:
    raw = normalize_value(value).strip()
    if not raw or raw in {"#NAME?", "#VALUE!", "#REF!", "XX"}:
        return ""
    m = re.search(r"-?\d+(?:\.\d+)?", raw.replace(",", ""))
    if not m:
        return ""
    try:
        return str(int(round(float(m.group(0)))))
    except Exception:
        return m.group(0)


def _read_step2_official_total_lots(excel_path: str | Path) -> str:
    """Return official Step 2 total lot count, excluding EXC rows.

    Preferred source is the visible/calculated total-lot cells in Step 2. If Excel
    has not cached formula results, compute from the Step 2 land-use column rather
    than Step 1. This preserves planner exclusions marked as EXC on Step 2.
    """
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception:
        return ""
    if STEP2_SHEET not in wb.sheetnames:
        return ""
    ws = wb[STEP2_SHEET]

    # Prefer a direct Step 2 table count when land-use rows are visible because it
    # respects planner EXC markings even when Excel formula caches are stale.
    # Stop after the ranked table has clearly ended.
    count = 0
    blank_streak = 0
    official_uses = {"SFR", "VAC", "MF", "MULTIFAMILY", "MULTI FAMILY", "COM", "COMMERCIAL"}
    for r in range(17, min(ws.max_row, 250) + 1):
        rank = ws.cell(r, 5).value
        desc = normalize_value(ws.cell(r, 1).value).strip()
        land = normalize_value(ws.cell(r, 2).value).strip().upper()
        size = ws.cell(r, 6).value
        has_rank = bool(_clean_number_text(rank))
        has_content = bool(desc or land or _clean_number_text(size))
        if not has_content:
            blank_streak += 1
            if blank_streak >= 5 and count:
                break
            continue
        blank_streak = 0
        if not has_rank and count:
            continue
        if land == "EXC":
            continue
        if land in official_uses:
            count += 1
    if count:
        return str(count)

    for addr in ("A15", "F11"):
        val = _clean_number_text(ws[addr].value)
        if val and val != "0":
            return val
    return ""


def _read_step2_official_minimum_lot_size(excel_path: str | Path) -> str:
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception:
        return ""
    if STEP2_SHEET not in wb.sheetnames:
        return ""
    ws = wb[STEP2_SHEET]
    n = lot_size_number(ws["E4"].value)
    if n and n != "0":
        return n
    # Formula cache fallback: first ranked size where cumulative area reaches 70%.
    for r in range(17, min(ws.max_row, 250) + 1):
        try:
            cum = float(ws.cell(r, 8).value)
        except Exception:
            continue
        if cum >= 0.70:
            n = lot_size_number(ws.cell(r, 6).value)
            if n and n != "0":
                return n
    return ""


def _read_step2_official_qualifying_values(excel_path: str | Path) -> dict[str, str]:
    """Count eligible Step 2 lots whose own area meets the proposed minimum.

    An existing Lots_Meeting_Minimum entry remains a planner override. Otherwise
    use the unranked property rows (B/D), including every tie at the minimum.
    The row where the ranked table reaches 70% is NOT the qualifying-lot count.
    Incomplete/uncalculated data keeps the visible placeholder without blocking
    document generation. This reader never changes the calculation workbook.
    """
    from decimal import Decimal, InvalidOperation

    unknown = {"Qualifying_Lot_Count": "X", "Qualifying_Lot_Count_Words": "X",
               "Qualifying_Lot_Count_Words_Cap": "X"}

    def numeric(value):
        try:
            result = Decimal(str(value).strip().replace(",", ""))
            return result if result.is_finite() else None
        except (InvalidOperation, ValueError):
            return None

    def formatted(count):
        words = number_to_words(count)
        return {"Qualifying_Lot_Count": str(count),
                "Qualifying_Lot_Count_Words": words,
                "Qualifying_Lot_Count_Words_Cap": words.capitalize()}

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception:
        return unknown
    try:
        if CASE_INFO_SHEET in wb.sheetnames:
            for row in wb[CASE_INFO_SHEET].iter_rows(min_col=1, max_col=2):
                if normalize_value(row[0].value).strip() != "Lots_Meeting_Minimum":
                    continue
                value = row[1].value
                if value is not None and str(value).strip():
                    count = numeric(value)
                    if count is None or count < 0 or count != count.to_integral_value():
                        return unknown
                    return formatted(int(count))
                break

        minimum = numeric(_read_step2_official_minimum_lot_size(excel_path))
        if minimum is None or minimum <= 0 or STEP2_SHEET not in wb.sheetnames:
            return unknown
        ws = wb[STEP2_SHEET]
        eligible_uses = {"SFR", "VAC", "MF", "MULTIFAMILY", "MULTI FAMILY", "COM", "COMMERCIAL"}
        count = total = 0
        for row in ws.iter_rows(min_row=17, max_row=min(ws.max_row, 250), max_col=4):
            description = normalize_value(row[0].value).strip()
            land_use = normalize_value(row[1].value).strip().upper()
            area = numeric(row[3].value)
            if land_use == "EXC":
                continue
            if not description and not land_use and area in (None, 0):
                continue
            if land_use not in eligible_uses or area is None or area < 0:
                return unknown
            total += 1
            if area >= minimum:
                count += 1
        # Do not publish a partial count when some formula results are missing.
        official_total = numeric(_read_step2_official_total_lots(excel_path))
        if not total or official_total != total:
            return unknown
        return formatted(count)
    finally:
        wb.close()


def _read_step2_official_case_calculations(excel_path: str | Path) -> dict[str, str]:
    values: dict[str, str] = {}
    support = _read_step2_visible_support_percentage(excel_path)
    if support:
        values["Support_Percentage"] = support
    sfr = _read_step2_visible_sfr_percentage(excel_path)
    if sfr:
        values["SFR_Percentage"] = sfr
    total = _read_step2_official_total_lots(excel_path)
    if total:
        values["Total_Lot_Count"] = total
        try:
            ti = int(total)
            values["Total_Lot_Count_Words"] = number_word(ti) if ti <= 50 else number_to_words(ti)
        except Exception:
            pass
    mls = _read_step2_official_minimum_lot_size(excel_path)
    if mls:
        values["Minimum_Lot_Size_Number"] = mls
        try:
            values["Minimum_Lot_Size_Words"] = number_to_words(int(mls.replace(",", "")))
        except Exception:
            pass
    values.update(_read_step2_official_qualifying_values(excel_path))
    return values


_prev_v137_calc_context = _build_calculation_context_v101

def _build_calculation_context_v101(excel_path: str | Path) -> dict[str, str]:
    context = _prev_v137_calc_context(excel_path)
    # Re-apply official Step 2 values last so Step 1/manual fallback cannot
    # override planner-visible calculation values.
    for k, v in _read_step2_official_case_calculations(excel_path).items():
        if v not in (None, ""):
            context[k] = normalize_value(v)
    # Case Info remains the source for planner-controlled narrative/review fields.
    try:
        case_info, _owners = load_case_and_owners(excel_path)
        for k in [
            "Application_Number", "Application_Type", "Street_Name", "Block_Description",
            "Side_Of_Street", "Cross_Streets", "Planner_Name", "Planner_Phone",
            "Planner_Email", "Council_Deadline", "Approval_Path", "Director_Name",
            "Director_Approval_Date", "PC_Signer", "PC_Approval_Date",
            "Neighborhood_Age_Statement", "Structure_Age_Statement", "Preservation_Finding",
            "Applicant_Name", "Applicant_Address_1", "Applicant_Address_2", "Applicant_City",
            "Applicant_State", "Applicant_Zip", "Sign_Count_Number", "Sign_Count_Words",
            "Sign_Placement", "Legal_File_No", "Ordinance_Year",
        ]:
            if case_info.get(k) not in (None, ""):
                context[k] = case_info[k]
    except Exception:
        pass
    # Keep derived text synchronized with the final official numbers.
    try:
        total = int(str(context.get("Total_Lot_Count", "")).replace(",", ""))
        context["Total_Lot_Count_Words"] = number_word(total) if total <= 50 else number_to_words(total)
    except Exception:
        pass
    try:
        q = int(str(context.get("Qualifying_Lot_Count", "")).replace(",", ""))
        context["Qualifying_Lot_Count_Words"] = number_word(q) if q <= 50 else number_to_words(q)
        context["Qualifying_Lot_Count_Words_Cap"] = number_word(q, capitalize=True) if q <= 50 else number_to_words(q).capitalize()
    except Exception:
        pass
    return context


def build_director_context(excel_path: str | Path) -> Dict[str, str]:
    return _build_calculation_context_v101(excel_path)


def build_pc_context(excel_path: str | Path) -> Dict[str, str]:
    case_info = build_director_context(excel_path)
    if not case_info.get("PC_Approval_Date"):
        case_info["PC_Approval_Date"] = ""
    if not case_info.get("PC_Signer"):
        case_info["PC_Signer"] = "Martha L. Stein, Chair"
    return case_info


def build_ordinance_context(excel_path: str | Path) -> Dict[str, str]:
    context = _build_calculation_context_v101(excel_path)
    area_upper, area_lower = _ordinance_area_from_case(context)
    context["Ordinance_Area_Upper"] = area_upper.upper()
    context["Ordinance_Area_Lower"] = area_lower
    _add_ordinance_template_fields(context)
    context.setdefault("Ordinance_Year", "2026")
    context.setdefault("Legal_File_No", "")
    context.setdefault("Director_Name", "Vonn Tran")
    return context

# --- v13.7b formula-cache fallback refinement -----------------------------------
# If Step 2 formula caches are blank because the workbook was edited outside Excel,
# still respect Step 2 EXC markings. Use Step 1 only to know how many source rows
# exist; never use Step 1 to override a visible Step 2 calculation.

def _step1_source_row_count_including_zero_area(excel_path: str | Path) -> int:
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception:
        return 0
    if STEP1_SHEET not in wb.sheetnames:
        return 0
    ws = wb[STEP1_SHEET]
    headers = [normalize_value(c.value) for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(headers) if h}
    addr_col = idx.get("Site_addr_1") or idx.get("Property Description") or 1
    tax_col = idx.get("TAX_ID") or 1
    count = 0
    blank_streak = 0
    for r in range(2, ws.max_row + 1):
        addr = _strip_null_text(ws.cell(r, addr_col).value)
        tax = _strip_null_text(ws.cell(r, tax_col).value)
        if addr or tax:
            count += 1
            blank_streak = 0
        else:
            blank_streak += 1
            if blank_streak >= 10 and count:
                break
    return count


def _step2_exc_rank_numbers(excel_path: str | Path) -> set[int]:
    out: set[int] = set()
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=False)
    except Exception:
        return out
    if STEP2_SHEET not in wb.sheetnames:
        return out
    ws = wb[STEP2_SHEET]
    for r in range(17, min(ws.max_row, 250) + 1):
        land = normalize_value(ws.cell(r, 2).value).strip().upper()
        if land != "EXC":
            continue
        rank_txt = _clean_number_text(ws.cell(r, 5).value)
        if rank_txt:
            try:
                out.add(int(rank_txt))
            except Exception:
                pass
    return out


def _read_step2_official_total_lots(excel_path: str | Path) -> str:
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception:
        return ""
    if STEP2_SHEET not in wb.sheetnames:
        return ""
    ws = wb[STEP2_SHEET]

    # 1) Best source when cached/visible: count visible Step 2 land-use rows.
    count = 0
    seen_any_visible_land_use = False
    official_uses = {"SFR", "VAC", "MF", "MULTIFAMILY", "MULTI FAMILY", "COM", "COMMERCIAL"}
    for r in range(17, min(ws.max_row, 250) + 1):
        land = normalize_value(ws.cell(r, 2).value).strip().upper()
        rank = _clean_number_text(ws.cell(r, 5).value)
        if not rank and count:
            break
        if land:
            seen_any_visible_land_use = True
        if land == "EXC":
            continue
        if land in official_uses:
            count += 1
    if count and seen_any_visible_land_use:
        return str(count)

    # 2) Next source: official total-lot cells.
    for addr in ("A15", "F11"):
        val = _clean_number_text(ws[addr].value)
        if val and val != "0":
            return val

    # 3) Formula-cache fallback: Step 1 source-row count minus Step 2 EXC rows.
    source_count = _step1_source_row_count_including_zero_area(excel_path)
    exc_count = len(_step2_exc_rank_numbers(excel_path))
    if source_count:
        return str(max(source_count - exc_count, 0))
    return ""


def _read_step2_visible_sfr_percentage(excel_path: str | Path) -> str:
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if STEP2_SHEET not in wb.sheetnames:
            return ""
        ws = wb[STEP2_SHEET]
        val = ws["H11"].value
        pct_val = pct(val)
        if pct_val:
            return pct_val
    except Exception:
        pass

    # Formula-cache fallback: count SFR rows from Step 1, but subtract Step 2 EXC ranks.
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception:
        return ""
    if STEP1_SHEET not in wb.sheetnames:
        return ""
    ws = wb[STEP1_SHEET]
    headers = [normalize_value(c.value) for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(headers) if h}
    addr_col = idx.get("Site_addr_1") or idx.get("Property Description") or 1
    tax_col = idx.get("TAX_ID") or 1
    land_col = idx.get("group_dscr") or idx.get("Land Use") or 1
    exc_ranks = _step2_exc_rank_numbers(excel_path)
    total = 0
    sfr = 0
    rank = 0
    blank_streak = 0
    for r in range(2, ws.max_row + 1):
        addr = _strip_null_text(ws.cell(r, addr_col).value)
        tax = _strip_null_text(ws.cell(r, tax_col).value)
        if not (addr or tax):
            blank_streak += 1
            if blank_streak >= 10 and total:
                break
            continue
        blank_streak = 0
        rank += 1
        if rank in exc_ranks:
            continue
        total += 1
        land = _strip_null_text(ws.cell(r, land_col).value).lower()
        if "sfr" in land or "single" in land:
            sfr += 1
    return str(int(round((sfr / total) * 100))) if total else ""

# --- v13.7c support formula-cache fallback --------------------------------------
# If the Step 2 support percent cell has no cached value, compute the same support
# calculation from Step 2's visible signed-in-support marks and Step 1 land area.
# This is only a formula-cache fallback; cached/visible Step 2 cells remain first.

def _read_step2_visible_support_percentage(excel_path: str | Path) -> str:
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if STEP2_SHEET in wb.sheetnames:
            ws = wb[STEP2_SHEET]
            for addr in ("H8", "G8", "F8", "H7", "G7", "F7"):
                val = _normalize_percent_text_for_template(ws[addr].value)
                if val:
                    try:
                        if 0 <= float(val) <= 100:
                            return val
                    except Exception:
                        pass
    except Exception:
        pass

    try:
        wb1 = openpyxl.load_workbook(excel_path, data_only=True)
        wb2 = openpyxl.load_workbook(excel_path, data_only=False)
    except Exception:
        return ""
    if STEP1_SHEET not in wb1.sheetnames or STEP2_SHEET not in wb2.sheetnames:
        return ""
    ws1 = wb1[STEP1_SHEET]
    ws2 = wb2[STEP2_SHEET]
    headers = [normalize_value(c.value) for c in ws1[1]]
    idx = {h: i + 1 for i, h in enumerate(headers) if h}
    addr_col = idx.get("Site_addr_1") or idx.get("Property Description") or 1
    tax_col = idx.get("TAX_ID") or 1
    area_col = idx.get("Total_Land_Area") or 1
    total_area = 0.0
    signed_area = 0.0
    rank = 0
    blank_streak = 0
    for r in range(2, ws1.max_row + 1):
        addr = _strip_null_text(ws1.cell(r, addr_col).value)
        tax = _strip_null_text(ws1.cell(r, tax_col).value)
        if not (addr or tax):
            blank_streak += 1
            if blank_streak >= 10 and total_area:
                break
            continue
        blank_streak = 0
        rank += 1
        try:
            area = float(str(ws1.cell(r, area_col).value).replace(",", ""))
        except Exception:
            area = 0.0
        total_area += area
        step2_row = 16 + rank
        signed = normalize_value(ws2.cell(step2_row, 3).value).strip().upper()
        if signed in {"Y", "YES", "TRUE", "1"}:
            signed_area += area
    if total_area > 0:
        return str(int(round((signed_area / total_area) * 100)))
    return ""

# Route-form context must use the same final Step 2/Case Info calculation context
# as approvals and ordinance, not the older direct workbook reader.
def build_routing_context(excel_path: str | Path) -> Dict[str, str]:
    case_info = _build_calculation_context_v101(excel_path)
    if not case_info.get("Approval_Path"):
        case_info["Approval_Path"] = "Director Approval"
    if not case_info.get("Council_Deadline"):
        case_info["Council_Deadline"] = ""
    return case_info

# --- v13.7d Case Info auto-read hardening --------------------------------------
# The Case Info tab mirrors Step 2 through formulas in B42:B48. When the workbook
# has not been recalculated by desktop Excel, openpyxl/data_only can see those
# formula cells as blank or stale cached values. The wizard should still show and
# use the official Step 2 values for the auto-read fields, so this final override
# reads Step 2 directly and applies those values after the normal Case Info merge.

_STEP2_AUTO_READ_FIELDS = {
    "Application_Number",
    "Application_Type",
    "Street_Name",
    "Lot_Block_Description",
    "Block_Description",
    "Side_Of_Street",
    "Cross_Streets",
    "Proposed_Min_Lot_Size",
    "Minimum_Lot_Size",
}


def _merge_nonblank_case_info(base: Dict[str, str], override: Dict[str, str]) -> Dict[str, str]:
    """Merge Case Info values without letting blank formula caches erase Step 2 defaults."""
    merged = dict(base)
    for key, value in override.items():
        if value not in (None, ""):
            merged[key] = value
    return merged


def load_case_and_owners(excel_path: str | Path) -> Tuple[Dict[str, str], List[Dict[str, str]]]:
    excel_path = Path(excel_path)
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    step2_defaults = read_step2_case_defaults(wb)
    case_info = _merge_nonblank_case_info(step2_defaults, read_case_info(wb))

    # Auto-read fields must always reflect Step 2, not stale/blank Case Info
    # formula caches. Planner-controlled fields still come from Case Info.
    for key in _STEP2_AUTO_READ_FIELDS:
        value = step2_defaults.get(key, "")
        if value not in (None, ""):
            case_info[key] = value

    if case_info.get("Application_Type"):
        case_info["Application_Type"] = normalize_application_type(case_info["Application_Type"])
    if not case_info.get("Block_Description"):
        case_info["Block_Description"] = case_info.get("Lot_Block_Description", "")
    if not case_info.get("Minimum_Lot_Size"):
        case_info["Minimum_Lot_Size"] = case_info.get("Proposed_Min_Lot_Size", "")

    addr2 = case_info.get("Applicant_Address_2", "")
    case_info["Applicant_Address_2_Block"] = ("\n" + addr2) if addr2 else ""
    case_info.setdefault("Sign_Count_Number", "2")
    case_info.setdefault("Sign_Count_Words", "two")
    case_info.setdefault("Sign_Placement", "one sign at each corner")
    if not case_info.get("Approval_Path"):
        case_info["Approval_Path"] = "Director Approval"
    owners = read_owner_rows_from_step1(wb)
    return case_info, owners
