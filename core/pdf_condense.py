"""Condense a multi-page attachment onto a single page.

HCAD property records are printed from Chrome onto letter paper, so a record
that is slightly too tall spills onto a second page that is mostly blank.
Managers want each HCAD record to read as one page in the RCA packet, so every
wizard runs its HCAD attachments through condense_pdf_to_one_page before the
packet is merged.

The approach is deliberately conservative: each page is trimmed to the box that
actually contains printed content, blank pages are dropped, and the trimmed
pages are stacked on one letter sheet, shrinking only as much as needed. Nothing
is cropped away, and the planner's original PDF is never modified.
"""

from __future__ import annotations

import io
import subprocess
import sys
from pathlib import Path

# Letter paper, in PDF points.
PAGE_WIDTH = 612.0
PAGE_HEIGHT = 792.0
MARGIN = 18.0
GAP = 9.0


def _get_fitz():
    """Return the PyMuPDF module, installing it once if it is missing.

    Mirrors the pypdf handling in the packet builders so the RCA Packet button
    still works on a machine where install_requirements.bat has not been run.
    """
    try:
        import fitz  # PyMuPDF
        return fitz
    except ImportError:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "PyMuPDF>=1.24.0"])
            import fitz  # type: ignore
            return fitz
        except Exception as exc:
            raise ModuleNotFoundError(
                "Condensing HCAD records needs the 'PyMuPDF' package. "
                "Run install_requirements.bat from the wizard folder, then try again. "
                f"Original error: {exc}"
            ) from exc


def _content_boxes(fitz, doc):
    """Return [(page_number, content_rect)] for every page that has content."""
    clips = []
    for page in doc:
        content = None
        page_area = page.rect.width * page.rect.height
        items = [(fitz.Rect(b[:4]), False) for b in page.get_text("blocks")]
        items.extend((d["rect"], True) for d in page.get_drawings())
        items.extend((fitz.Rect(i["bbox"]), False) for i in page.get_image_info())
        for box, is_drawing in items:
            box = box & page.rect
            if box.is_empty:
                continue
            # HCAD paints a page-sized background rectangle. Counting it as
            # content would keep the blank bottom of every page.
            if is_drawing and box.width * box.height >= 0.9 * page_area:
                continue
            content = box if content is None else content | box
        if content is None or content.width < 1 or content.height < 1:
            continue  # blank page
        # Include stroke widths and glyph overshoot at the content edges.
        content = (content + (-2, -2, 2, 2)) & page.rect
        clips.append((page.number, content))
    return clips


def condense_pdf_to_one_page(pdf_path: str | Path, output_pdf: str | Path) -> Path:
    """Return a one-page version of pdf_path, written to output_pdf.

    Returns pdf_path unchanged when the file is already one page, or when the
    PDF cannot be condensed, so a packet build never fails over this step.
    """
    pdf_path = Path(pdf_path)
    try:
        fitz = _get_fitz()
        src = fitz.open(str(pdf_path))
        try:
            if src.page_count <= 1:
                return pdf_path

            # Text/drawing bounds use unrotated coordinates, while page.rect
            # includes /Rotate. Bake that display rotation into this in-memory
            # copy before cropping or embedding; the original file is untouched.
            if any(page.rotation for page in src):
                from pypdf import PdfReader, PdfWriter
                reader = PdfReader(io.BytesIO(pdf_path.read_bytes()))
                writer = PdfWriter()
                for page in reader.pages:
                    if page.rotation:
                        # Transform the crop/media boxes together with content.
                        page.transfer_rotation_to_content()
                    writer.add_page(page)
                normalized = io.BytesIO()
                writer.write(normalized)
                src.close()
                src = fitz.open(stream=normalized.getvalue(), filetype='pdf')
            clips = _content_boxes(fitz, src)
            if not clips:
                return pdf_path
            avail_w = PAGE_WIDTH - 2 * MARGIN
            avail_h = PAGE_HEIGHT - 2 * MARGIN - GAP * (len(clips) - 1)
            if avail_h <= 0:
                return pdf_path
            widest = max(rect.width for _, rect in clips)
            total_h = sum(rect.height for _, rect in clips)
            # A single remaining landscape page must fit too (for example,
            # when its second page was blank). Never enlarge the source.
            scale = min(avail_w / widest, avail_h / total_h, 1.0)

            out = fitz.open()
            new_page = out.new_page(width=PAGE_WIDTH, height=PAGE_HEIGHT)
            y = MARGIN
            for page_number, rect in clips:
                w, h = rect.width * scale, rect.height * scale
                x = (PAGE_WIDTH - w) / 2
                new_page.show_pdf_page(fitz.Rect(x, y, x + w, y + h), src, page_number, clip=rect)
                y += h + GAP

            output_pdf = Path(output_pdf)
            output_pdf.parent.mkdir(parents=True, exist_ok=True)
            out.save(str(output_pdf), garbage=3, deflate=True)
            out.close()
            return output_pdf
        finally:
            src.close()
    except Exception:
        return pdf_path


def condense_hcad_pdfs(hcad_pdfs, work_dir: str | Path) -> list:
    """Condense each HCAD record PDF to one page, in packet order."""
    work_dir = Path(work_dir)
    condensed = []
    for index, pdf in enumerate(hcad_pdfs, start=1):
        pdf = Path(pdf)
        target = work_dir / f"hcad_onepage_{index:02d}_{pdf.stem}.pdf"
        condensed.append(condense_pdf_to_one_page(pdf, target))
    return condensed
