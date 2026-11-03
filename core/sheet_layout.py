"""Page-layout helpers for exporting a Step 2 calculation sheet to PDF.

The calculation sheets store the row heights Excel used on the planner's screen.
LibreOffice does the PDF conversion for the RCA packet and measures text
slightly wider, so a property description that fits two lines in Excel can wrap
to three and be cut off at the bottom of a fixed-height row. That silently drops
data a reviewer needs, so every wizard grows its row heights before converting.

Heights are only ever increased. The official form keeps its spacing, and rows
the planner deliberately made tall are left alone.
"""

from __future__ import annotations

# Squeeze a sheet onto one page only while the text stays at least this
# fraction of its designed size. Below it, letting the table run onto a second
# page is more readable than shrinking further.
MIN_ONE_PAGE_SCALE = 0.6

DEFAULT_COLUMN_WIDTH = 8.43  # Excel's default, in characters
DEFAULT_ROW_HEIGHT = 15.0  # points
LETTER_WIDTH = 612.0
LETTER_HEIGHT = 792.0


def wrapped_line_count(text, width_chars: int) -> int:
    """Greedy word-wrap line count for a cell rendered width_chars wide."""
    if width_chars < 1:
        return 1
    lines, used = 1, 0
    for word in str(text).split():
        needed = len(word) if used == 0 else len(word) + 1
        if used + needed <= width_chars:
            used += needed
            continue
        lines += 1
        used = len(word)
        if used > width_chars:  # a single word longer than the column
            lines += (used - 1) // width_chars
            used = (used - 1) % width_chars + 1
    return lines


def _column_width(ws, col: int) -> float:
    from openpyxl.utils import get_column_letter

    dim = ws.column_dimensions.get(get_column_letter(col))
    return getattr(dim, "width", None) or DEFAULT_COLUMN_WIDTH


def _merge_maps(ws):
    """Return (rows spanned by multi-row merges, anchor -> column span, covered cells)."""
    multi_row_rows = set()
    spans = {}
    covered = set()
    for rng in ws.merged_cells.ranges:
        if rng.max_row > rng.min_row:
            multi_row_rows.update(range(rng.min_row, rng.max_row + 1))
        spans[(rng.min_row, rng.min_col)] = (rng.min_col, rng.max_col)
        for row in range(rng.min_row, rng.max_row + 1):
            for col in range(rng.min_col, rng.max_col + 1):
                if (row, col) != (rng.min_row, rng.min_col):
                    covered.add((row, col))
    return multi_row_rows, spans, covered


def autofit_row_heights(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    """Grow row heights so wrapped text is not clipped in the exported PDF."""
    multi_row_rows, spans, covered = _merge_maps(ws)

    for row in range(min_row, max_row + 1):
        # A cell merged across rows has no single row to grow, so leave it alone.
        if row in multi_row_rows:
            continue
        needed = 0.0
        for col in range(min_col, max_col + 1):
            if (row, col) in covered:
                continue
            cell = ws.cell(row, col)
            if cell.value in (None, "") or not cell.alignment.wrap_text:
                continue
            first_col, last_col = spans.get((row, col), (col, col))
            width = sum(_column_width(ws, c) for c in range(first_col, last_col + 1))
            font_size = cell.font.size or 11
            # Column width is measured in characters of the default 11pt font.
            chars = max(1, int(width * 11 / font_size))
            lines = wrapped_line_count(cell.value, chars)
            needed = max(needed, lines * font_size * 1.3 + 3)
        if needed <= 0:
            continue
        current = getattr(ws.row_dimensions.get(row), "height", None) or 0
        if needed > current:
            ws.row_dimensions[row].height = needed


def one_page_scale(ws, min_row: int, max_row: int, min_col: int, max_col: int,
                   margin_in: float = 0.5) -> float:
    """Estimate the shrink factor needed to fit the print range on one page.

    Returns the scale the converter would have to apply, where 1.0 means it
    already fits.
    """
    usable_w = LETTER_WIDTH - margin_in * 72
    usable_h = LETTER_HEIGHT - margin_in * 72

    width_px = 0.0
    for col in range(min_col, max_col + 1):
        width_px += _column_width(ws, col) * 7 + 5  # Excel character-width-to-pixel formula
    content_w = width_px * 0.75  # 96 dpi pixels to points

    default_height = ws.sheet_format.defaultRowHeight or DEFAULT_ROW_HEIGHT
    content_h = 0.0
    for row in range(min_row, max_row + 1):
        content_h += getattr(ws.row_dimensions.get(row), "height", None) or default_height

    if content_w <= 0 or content_h <= 0:
        return 1.0
    return min(1.0, usable_w / content_w, usable_h / content_h)


def apply_staff_analysis_page_setup(ws, print_area: str, min_row: int, max_row: int,
                                    min_col: int, max_col: int, header_row: int | None = None) -> None:
    """Set the print area and page fitting for a Step 2 staff-analysis export.

    Rows are grown first so nothing is clipped, then the sheet is squeezed onto
    one page when that still leaves it readable. When the table genuinely needs a
    second page, header_row is repeated at the top of it so the continued
    columns stay labeled.
    """
    autofit_row_heights(ws, min_row, max_row, min_col, max_col)

    ws.print_area = print_area
    ws.page_setup.fitToWidth = 1
    fits_one_page = one_page_scale(ws, min_row, max_row, min_col, max_col) >= MIN_ONE_PAGE_SCALE
    ws.page_setup.fitToHeight = 1 if fits_one_page else 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    if header_row and not fits_one_page:
        ws.print_title_rows = f"{header_row}:{header_row}"
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25
