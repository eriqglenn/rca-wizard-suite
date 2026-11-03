"""Workbook formula-cache repair for the RCA Wizard Suite.

Some calculation workbooks are written by tools that stamp every formula with
an EMPTY cached value (``<f>...</f><v></v>``) and omit the workbook recalc
flag.  Excel then treats the empty cache as the current result and shows
Step 2 as blank even after Ctrl+Alt+F9, and openpyxl's ``data_only=True``
read returns ``None`` for every formula.  The formulas and source data are
fine; only the saved cache is poisoned.

This module provides two things:

  needs_recalc(path)   - cheap detector: True if the Step 2 formula cache looks
                         empty/stale (poisoned), so callers can warn or repair.

  recalculated_copy()  - returns a path to a recalculated copy of the workbook
                         (Excel, with LibreOffice fallback), or the original path if recalc
                         is unavailable.  Never raises; callers read from the
                         returned path.

  repair_in_place_copy() - writes a *fixed* .xlsx next to a chosen output that
                           strips the empty cache and sets fullCalcOnLoad, so
                           opening it in Excel recalculates immediately.

Reading still prefers the workbook's own cached values when they are present
(fast, no external process); recalc is only triggered when the cache is empty.
"""

from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from .soffice import find_soffice  # resolver already used elsewhere in the suite


def _step2_sheet_xml(path: str | Path) -> str | None:
    """Return the Step 2 worksheet XML, or None if it can't be located."""
    try:
        z = zipfile.ZipFile(str(path))
        wb = z.read("xl/workbook.xml").decode("utf-8", "replace")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
        # Relationship Id -> Target, handling either attribute order.
        relmap = dict(re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', rels))
        for rid, tgt in re.findall(r'Target="([^"]+)"[^>]*?Id="([^"]+)"', rels):
            relmap.setdefault(tgt, rid)  # harmless; primary map already built
        for tag in re.findall(r"<sheet\b[^>]*/>", wb):
            name_m = re.search(r'name="([^"]+)"', tag)
            rid_m = re.search(r'r:id="([^"]+)"', tag)
            if not name_m or not rid_m:
                continue
            if name_m.group(1).lower().startswith("step 2"):
                target = relmap.get(rid_m.group(1), "")
                if not target:
                    continue
                target = target.lstrip("/")
                if not target.startswith("xl/"):
                    target = "xl/" + target
                return z.read(target).decode("utf-8", "replace")
    except Exception:
        return None
    return None


def needs_recalc(path: str | Path) -> bool:
    """True if Step 2 formulas appear to have an empty/poisoned cache.

    Heuristic: among Step 2 formula cells, if the majority store an empty
    ``<v></v>``/``<v/>`` (or no value at all), the cache is stale.
    """
    xml = _step2_sheet_xml(path)
    if not xml:
        return False
    try:
        root=ET.fromstring(xml)
    except ET.ParseError:
        return False
    ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    cells=[c for c in root.findall('.//s:c',ns) if c.find('s:f',ns) is not None]
    if not cells:return False
    # Shared-formula followers use <f .../> rather than </f>. Counting closing
    # tags missed their cleared caches in newly imported City workbooks.
    empty=sum(c.find('s:v',ns) is None or not (c.find('s:v',ns).text or '').strip() for c in cells)
    return empty >= max(1,len(cells)//2)


def recalculated_copy(path: str | Path, timeout: int = 120) -> Path:
    """Return a path to a recalculated copy of the workbook.

    If recalculation is not needed, or LibreOffice is unavailable, returns the
    original path unchanged.  Never raises.
    """
    src = Path(path)
    try:
        if not needs_recalc(src):
            return src
        soffice = find_soffice()
        tmpdir = Path(tempfile.mkdtemp(prefix="rca_recalc_"))
        input_dir=tmpdir/'input';input_dir.mkdir()
        work = input_dir / ("_src_" + src.name if src.suffix else src.name + ".xlsx")
        # Strip the poisoned empty cache and set fullCalcOnLoad BEFORE handing the
        # file to LibreOffice; otherwise the convert preserves the empty <v> cells
        # and the recalculated copy still reads blank.
        repaired = repair_in_place_copy(src, work)
        if repaired is None:
            shutil.copy2(src, work)
        # Prefer Excel for the City's native formulas and embedded artwork.
        # LibreOffice remains available on computers without desktop Excel.
        from .office_convert import recalc_with_excel
        got = recalc_with_excel(work, tmpdir / (work.stem + "_excel.xlsx"))
        if got:return got
        if not soffice:return src
        # Converting to xlsx forces LibreOffice to recalculate the formulas.
        # Prefer an isolated user profile so a separately running LibreOffice
        # instance cannot turn this convert into a silent no-op; if that form
        # fails for any reason, fall back to a plain convert.
        # Conversion output must differ from the staged input. Otherwise the
        # existing input file can be mistaken for a successful recalculation.
        out = tmpdir / (work.stem + ".xlsx")
        profile = (tmpdir / "lo_profile").as_uri()
        attempts = [
            [soffice, "-env:UserInstallation=" + profile, "--headless",
             "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
             "--outdir", str(tmpdir), str(work)],
            [soffice, "--headless", "--calc",
             "--convert-to", "xlsx:Calc MS Excel 2007 XML",
             "--outdir", str(tmpdir), str(work)],
        ]
        for cmd in attempts:
            try:
                subprocess.run(cmd, timeout=timeout, capture_output=True)
            except Exception:
                continue
            if out.exists():
                break
        return out if out.exists() else (work if repaired is not None else src)
    except Exception:
        return src


def repair_in_place_copy(path: str | Path, dest: str | Path) -> Path | None:
    """Write a repaired .xlsx to ``dest``: strip the empty formula cache and
    set ``fullCalcOnLoad`` so Excel recalculates on open.  Returns the dest
    path on success, or None on failure.  Does not modify the original."""
    try:
        src = Path(path)
        dest = Path(dest)
        zin = zipfile.ZipFile(str(src))
        with zipfile.ZipFile(str(dest), "w", zipfile.ZIP_DEFLATED) as zout:
            for n in zin.namelist():
                data = zin.read(n)
                if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"):
                    t = data.decode("utf-8", "replace")
                    t = re.sub(r"(</f>)<v\s*/>", r"\1", t)
                    t = re.sub(r"(</f>)<v></v>", r"\1", t)
                    data = t.encode("utf-8")
                elif n == "xl/workbook.xml":
                    t = data.decode("utf-8", "replace")
                    if "<calcPr" in t:
                        t = re.sub(r"<calcPr[^>]*/>",
                                   '<calcPr calcId="191028" fullCalcOnLoad="1"/>', t)
                    else:
                        t = t.replace("</workbook>",
                                      '<calcPr calcId="191028" fullCalcOnLoad="1"/></workbook>')
                    data = t.encode("utf-8")
                zout.writestr(n, data)
        return dest
    except Exception:
        return None
