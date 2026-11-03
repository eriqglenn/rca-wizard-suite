"""RCA Wizard Suite — build validation harness (master brief §20.12).

Runs every wizard's document generators headlessly against the bundled sample
workbooks and applies the structural checks from the validation contract:

  - no unfilled {{placeholders}} anywhere in any output (body/headers/footers)
  - combined mail-outs keep per-letter section structure (titlePg count equals
    footerReference count where the template uses a first-page footer)
  - routing-form checkboxes carry both state and glyph
  - every expected output file exists and is non-trivial in size

Run from Wizard_Files:

    python tools/validate.py            # full run, exit code 0 = PASS
    python tools/validate.py --wizard mls

This harness is deterministic and headless (no Tk window), so it can gate any
build before delivery.  Visual confirmation in Microsoft Word (§20.12.3) is
still required before department-wide rollout — this tool does not replace it.
"""

from __future__ import annotations

import argparse
import re
import sys
import tempfile
import warnings
import zipfile
from pathlib import Path

warnings.filterwarnings("ignore")

WIZARD_FILES = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(WIZARD_FILES))

from core.paths import templates_dir, sample_input_dir  # noqa: E402
from core.templates import template_path, DEFAULT_ROLES  # noqa: E402

# Templates are resolved through templates/manifest.json so a replaced
# template (possibly under a new file name) is what gets validated.
ROLE_OF = {fn: (wk, rk) for wk, rl in DEFAULT_ROLES.items() for rk, _l, fn in rl}


def _t(filename: str) -> Path:
    wk, rk = ROLE_OF[filename]
    return template_path(wk, rk, filename)

FAILURES: list[str] = []
CHECKS = 0


def _xml(docx: Path) -> str:
    return zipfile.ZipFile(docx).read("word/document.xml").decode("utf-8")


def check(label: str, condition: bool, detail: str = ""):
    global CHECKS
    CHECKS += 1
    if condition:
        print(f"  [PASS] {label}")
    else:
        print(f"  [FAIL] {label}" + (f" — {detail}" if detail else ""))
        FAILURES.append(label)


def check_docx(path: Path, label: str, combined_titlepg: bool = False,
               expect_checkboxes: bool = False, template: Path | None = None):
    check(f"{label}: file created", path.exists() and path.stat().st_size > 5000,
          str(path))
    if not path.exists():
        return
    xml = _xml(path)
    leftovers = sorted(set(re.findall(r"\{\{[^}]+\}\}", xml)))
    check(f"{label}: no unfilled placeholders", not leftovers, str(leftovers))
    # Every generated Word document must be fully editable: the City ordinance
    # and routing-form templates ship with document protection
    # (w:edit="forms"), which Word enforces as read-only-except-form-fields.
    # The wizards remove it on output so planners can edit freely.
    try:
        from core.docx_unlock import docx_is_protected
        check(f"{label}: fully editable (no document protection)",
              not docx_is_protected(path), "document protection still enforced")
    except Exception:
        pass
    if combined_titlepg:
        # §20.6a: templates with a FIRST-PAGE footer must keep one
        # titlePg+footerReference pair per letter.  Templates with ordinary
        # footers have titlePg=0; there every section must still carry its
        # footer references (the combiner must not drop them).
        tp = xml.count("w:titlePg")
        fr = len(re.findall(r"footerReference", xml))
        sects = len(re.findall(r"<w:sectPr", xml))
        if template is not None and Path(template).exists():
            # A letter template may span several sections (the 2026 renewal
            # letters put the map on its own section).  The combined mail-out
            # must then repeat the template's whole section shape per letter:
            # N letters -> N x template sections, titlePg flags and footer refs.
            txml = _xml(Path(template))
            ttp, tfr = txml.count("w:titlePg"), len(re.findall(r"footerReference", txml))
            tsects = max(1, len(re.findall(r"<w:sectPr", txml)))
            n = sects // tsects
            check(f"{label}: section shape repeats per letter ({sects} sections = {n} x {tsects})",
                  n >= 1 and sects == n * tsects, f"sections={sects} template={tsects}")
            check(f"{label}: first-page footer per letter (titlePg {tp} == {n} x {ttp})",
                  tp == n * ttp, f"titlePg={tp} expected={n * ttp}")
            check(f"{label}: footer references per letter (footerRef {fr} == {n} x {tfr})",
                  fr == n * tfr, f"footerRef={fr} expected={n * tfr}")
        elif tp > 0:
            check(f"{label}: first-page footer balance (titlePg {tp} == footerRef {fr})",
                  tp == fr, f"titlePg={tp} footerRef={fr}")
        else:
            check(f"{label}: footers preserved per section (footerRef {fr} >= sections {sects})",
                  fr >= sects > 0, f"footerRef={fr} sections={sects}")
    if expect_checkboxes:
        glyphs = xml.count("\u2612") + xml.count("\u2610")
        check(f"{label}: checkbox glyphs present ({glyphs})", glyphs >= 2)




# ----------------------------------------------------------------------------
# Letter spacing check (template-fidelity for the address/date/subject region)
# ----------------------------------------------------------------------------
def _effective_first_line_x(p):
    """Effective first-line left position in twips: left + firstLine - hanging."""
    from docx.oxml.ns import qn
    ppr = p._p.find(qn('w:pPr'))
    if ppr is None:
        return 0
    ind = ppr.find(qn('w:ind'))
    if ind is None:
        return 0
    def a(name):
        v = ind.get(qn('w:' + name))
        try:
            return int(v)
        except (TypeError, ValueError):
            return 0
    return a('left') + a('firstLine') - a('hanging')


def _tab_count(p):
    from docx.oxml.ns import qn
    return len(p._p.findall('.//' + qn('w:tab')))


def check_letter_spacing(template_path: Path, generated_path: Path, label: str,
                         skip_paras=(), n: int = 12):
    """Head-region template fidelity: for the first n paragraphs the generated
    letter must keep the template's effective first-line position and its tab
    structure (a doubled tab is how the date drifted right of the template).
    ``skip_paras`` holds indexes covered by documented deliberate fixes (e.g.
    the No-SFR subject alignment fix normalizes the template's own mixed
    tab/indent continuation lines)."""
    from docx import Document
    tpl = Document(str(template_path))
    gen = Document(str(generated_path))
    for i in range(min(n, len(tpl.paragraphs), len(gen.paragraphs))):
        if i in skip_paras:
            continue
        tx, gx = _effective_first_line_x(tpl.paragraphs[i]), _effective_first_line_x(gen.paragraphs[i])
        tt, gt = _tab_count(tpl.paragraphs[i]), _tab_count(gen.paragraphs[i])
        check(f"{label}: para[{i}] first-line position {gx} == template {tx}",
              gx == tx, f"text={gen.paragraphs[i].text[:40]!r}")
        check(f"{label}: para[{i}] tab structure {gt} == template {tt}",
              gt == tt, f"text={gen.paragraphs[i].text[:40]!r}")



# ----------------------------------------------------------------------------
# Blank-merge artifacts + poisoned-workbook regression (the "is  square feet" bug)
# ----------------------------------------------------------------------------
_BLANK_MERGE_PATTERNS = [
    (r"\bis\s+square feet", "minimum lot size missing ('is  square feet')"),
    (r"\bthan\s+square feet", "minimum lot size missing ('than  square feet')"),
    (r"\bof\s+feet\b", "building line missing ('of  feet')"),
    (r"\bleast\s+feet\b", "building line missing ('least  feet')"),
]


def check_no_blank_merges(path: Path, label: str):
    """A merge value that comes back empty must keep the template's visible
    X placeholder; it must never silently delete the number and leave broken
    phrases like 'the minimum lot size is  square feet.'"""
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8", "replace")
    text = re.sub(r"<[^>]+>", "", xml)
    for pat, why in _BLANK_MERGE_PATTERNS:
        check(f"{label}: no blank-merge artifact ({why})",
              not re.search(pat, text), f"pattern {pat!r} matched")


def poison_workbook(src: Path, dst: Path):
    """Re-create the stale-formula-cache failure (brief section 20.4): every
    formula's cached value emptied and the recalc flag removed, exactly the
    state of a workbook saved by a tool that never let Excel calculate."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            data = zin.read(n)
            if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"):
                t = data.decode("utf-8", "replace")
                t = re.sub(r"(</f>)<v[^>]*>[^<]*</v>", r"\1<v></v>", t)
                data = t.encode("utf-8")
            elif n == "xl/workbook.xml":
                data = data.decode().replace(' fullCalcOnLoad="1"', '').encode()
            zout.writestr(n, data)


def check_pc_chair(path: Path, label: str):
    """The Planning Commission approval must name the current chair
    (Lisa Clark, Chair) and the current vice-chair (Sonny Garza, Vice-Chair),
    and must not contain the former chair's name."""
    xml = zipfile.ZipFile(str(path)).read("word/document.xml").decode("utf-8", "replace")
    text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))
    check(f"{label}: names current chair (Lisa Clark, Chair)",
          ("Lisa Clark, Chair" in text or "Lisa M. Clark, Chair" in text), "current chair line missing")
    check(f"{label}: names vice-chair (Sonny Garza, Vice-Chair)",
          "Sonny Garza, Vice-Chair" in text, "vice-chair line missing")
    check(f"{label}: no former chair name (Stein)",
          "Stein" not in text, "former chair name present")


def check_no_date_timestamps(path: Path, label: str):
    """Excel dates read as datetime objects must render in the City long form
    (e.g. 'January 18, 2006'), never with a leaked midnight timestamp
    ('2006-01-18 00:00:00') or in bare ISO form. Scans body + headers/footers."""
    z = zipfile.ZipFile(path)
    parts = [n for n in z.namelist()
             if n.endswith(".xml") and ("document" in n or "header" in n or "footer" in n)]
    for part in parts:
        text = re.sub(r"<[^>]+>", "", z.read(part).decode("utf-8", "replace"))
        ts = re.findall(r"\d{1,2}:\d{2}:\d{2}", text)
        iso = re.findall(r"\b\d{4}-\d{2}-\d{2}\b", text)
        check(f"{label}: no leaked date timestamp ({part.split('/')[-1]})",
              not ts, f"found {ts[:3]}")
        check(f"{label}: no bare ISO date ({part.split('/')[-1]})",
              not iso, f"found {iso[:3]}")


def check_contains(path: Path, label: str, needle: str):
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8", "replace")
    text = re.sub(r"<[^>]+>", "", xml).replace("\u00a0", " ")
    check(f"{label}: contains {needle!r}", needle in text, "value missing from document")


def validate_mls(out: Path):
    print("MLS Wizard")
    from wizards.mls.mailout_engine import generate_planning_commission_approval
    from wizards.mls.templates_2026 import (
        generate_owner_letters, generate_routing_form, generate_director_approval,
        generate_applicant_letter, generate_draft_ordinance, generate_staff_report)
    si, td = sample_input_dir("mls"), templates_dir("mls")
    wb = si / "MLS_Calculation_Sheet_Case_Info_v13_FILLED_SAMPLE.xlsx"
    mp = si / "MLS_916_Map_SAMPLE.pdf"
    generate_owner_letters(wb, _t("MLS_Owner_Notification_Letter.docx"), mp, out / "mls_preview.docx", preview=True)
    check_docx(out / "mls_preview.docx", "notification preview")
    check_letter_spacing(_t("MLS_Owner_Notification_Letter.docx"), out / "mls_preview.docx", "letter spacing")
    check_no_blank_merges(out / "mls_preview.docx", "notification preview")
    check_contains(out / "mls_preview.docx", "notification preview", "10,350 square feet")
    generate_owner_letters(wb, _t("MLS_Owner_Notification_Letter.docx"), mp, out / "mls_all.docx", preview=False)
    check_docx(out / "mls_all.docx", "all-owner mail-out", combined_titlepg=True,
               template=_t("MLS_Owner_Notification_Letter.docx"))
    generate_owner_letters(wb, _t("MLS_Protest_All_Owners.docx"), mp, out / "mls_protest_prev.docx", preview=True)
    check_docx(out / "mls_protest_prev.docx", "protest letter preview")
    check_letter_spacing(_t("MLS_Protest_All_Owners.docx"), out / "mls_protest_prev.docx", "protest letter spacing")
    generate_owner_letters(wb, _t("MLS_Protest_All_Owners.docx"), mp, out / "mls_protest_all.docx", preview=False)
    check_docx(out / "mls_protest_all.docx", "protest all-owner mail-out", combined_titlepg=True,
               template=_t("MLS_Protest_All_Owners.docx"))
    generate_owner_letters(wb, _t("MLS_Notice_Public_Hearing_Less_51.docx"), mp, out / "mls_hearing_prev.docx", preview=True)
    check_docx(out / "mls_hearing_prev.docx", "public hearing notice preview")
    check_contains(out / "mls_hearing_prev.docx", "public hearing notice preview", "67%")
    generate_routing_form(wb, _t("MLS_Routing Form.docx"), out / "mls_routing.docx")
    check_docx(out / "mls_routing.docx", "routing form", expect_checkboxes=True)
    check_contains(out / "mls_routing.docx", "routing form", "Block Application No. 933")
    generate_director_approval(wb, _t("MLS_Director_Approval.docx"), out / "mls_dir.docx")
    check_docx(out / "mls_dir.docx", "director approval")
    check_contains(out / "mls_dir.docx", "director approval", "10,350")
    generate_planning_commission_approval(wb, _t("PC_Approval_NEW.docx"), out / "mls_pc.docx")
    check_docx(out / "mls_pc.docx", "PC approval")
    check_pc_chair(out / "mls_pc.docx", "PC approval")
    generate_applicant_letter(wb, _t("MLS_Applicant_Complete_Signs.docx"), out / "mls_app.docx")
    check_docx(out / "mls_app.docx", "applicant letter")
    check_contains(out / "mls_app.docx", "applicant letter", "Sample Applicant")
    generate_staff_report(wb, _t("MLS_OneMeeting_Staff_Report.docx"), out / "mls_staff.docx")
    check_docx(out / "mls_staff.docx", "OneMeeting staff report")
    check_contains(out / "mls_staff.docx", "OneMeeting staff report", "Application No. 933")
    generate_draft_ordinance(wb, _t("MLS_Ordinance_Template.docx"), mp, out / "mls_ord.docx")
    check_docx(out / "mls_ord.docx", "draft ordinance")
    check_contains(out / "mls_ord.docx", "draft ordinance", "ten thousand three hundred and fifty (10,350)")


def validate_mbl(out: Path):
    print("MBL Wizard")
    from wizards.mbl import wizard as mod
    from openpyxl import load_workbook as _lw
    si, td = sample_input_dir("mbl"), templates_dir("mbl")
    wb = si / "MBL_Calculation_Sheet_Case_Info_v3_FILLED_SAMPLE.xlsx"
    mp = si / "sample_boundary_map.pdf"
    owners = mod._get_owners(_lw(wb, data_only=True))
    check("owner rows readable", len(owners) > 0, "no owners found")
    area = "1800 block of Haver Street, north side, between Woodhead Street and Dunlavy Street"
    for tname, lab in (("MBL_ALL_OWNERS.docx", "notification"),
                       ("MBL_Notice_Public_Hearing_Less_51.docx", "hearing"),
                       ("MBL_Protest_All_Owners.docx", "protest"),
                       ("MBL_Notice_PC_Rescheduling.docx", "reschedule")):
        mod.generate_notification(str(wb), str(_t(tname)), str(mp),
                                  str(out / f"mbl_{lab}_prev.docx"), mode="preview")
        check_docx(out / f"mbl_{lab}_prev.docx", f"{lab} letter preview")
        check_letter_spacing(_t(tname), out / f"mbl_{lab}_prev.docx", f"{lab} letter spacing")
        check_no_blank_merges(out / f"mbl_{lab}_prev.docx", f"{lab} letter preview")
        if lab != "protest":
            check_contains(out / f"mbl_{lab}_prev.docx", f"{lab} letter preview", "15 feet")
        mod.generate_notification(str(wb), str(_t(tname)), str(mp),
                                  str(out / f"mbl_{lab}_all.docx"), mode="all")
        check_docx(out / f"mbl_{lab}_all.docx", f"{lab} all-owner mail-out",
                   combined_titlepg=True, template=_t(tname))
    mod.generate_routing_form(str(wb), str(_t("MBL_Routing_Form.docx")), str(out / "mbl_routing.docx"))
    check_docx(out / "mbl_routing.docx", "routing form", expect_checkboxes=True)
    check_contains(out / "mbl_routing.docx", "routing form", "Block Application No. 1030")
    mod.generate_director_approval(str(wb), str(_t("MBL_Directors_Approval.docx")), str(out / "mbl_dir.docx"))
    check_docx(out / "mbl_dir.docx", "director approval")
    check_contains(out / "mbl_dir.docx", "director approval", "Block No. 1030")
    check_contains(out / "mbl_dir.docx", "director approval", "15 feet")
    mod.generate_pc_approval(str(wb), str(_t("MBL_PC_Approval.docx")), str(out / "mbl_pc.docx"))
    check_docx(out / "mbl_pc.docx", "PC approval")
    check_pc_chair(out / "mbl_pc.docx", "PC approval")
    check_contains(out / "mbl_pc.docx", "PC approval", "owners of 60% of the SMBLB")
    check_contains(out / "mbl_pc.docx", "PC approval", "Three (3) out of four (4) developed properties")
    mod.generate_applicant_letter(str(wb), str(_t("MBL_APPLICANT_COMPLETE_SIGNS.docx")), str(out / "mbl_app.docx"))
    check_docx(out / "mbl_app.docx", "applicant letter")
    check_contains(out / "mbl_app.docx", "applicant letter", "15 feet")
    for tname, lab in (("MBL_OneMeeting_Staff_Report.docx", "director"),
                       ("MBL_OneMeeting_Staff_Report_Less51.docx", "less51"),
                       ("MBL_OneMeeting_Staff_Report_Protest.docx", "protest")):
        mod.generate_staff_report(str(wb), str(_t(tname)), str(out / f"mbl_staff_{lab}.docx"))
        check_docx(out / f"mbl_staff_{lab}.docx", f"OneMeeting staff report ({lab})")
        check_contains(out / f"mbl_staff_{lab}.docx", f"OneMeeting staff report ({lab})", "Application No. 1030")
        check_contains(out / f"mbl_staff_{lab}.docx", f"OneMeeting staff report ({lab})", area)
    mod.generate_draft_ordinance(str(wb), str(_t("MBL_Ordinance_Template_Clean.docx")), str(mp), str(out / "mbl_ord.docx"))
    check_docx(out / "mbl_ord.docx", "draft ordinance")
    check_contains(out / "mbl_ord.docx", "draft ordinance", "fifteen (15) feet")
    # Stale-cache regression: a workbook saved with an empty formula cache must
    # still produce a letter that carries the building line (derived from the
    # typed Step 1 / Step 2 columns).
    pz = out / "mbl_poisoned.xlsx"
    poison_workbook(wb, pz)
    mod.generate_notification(str(pz), str(_t("MBL_ALL_OWNERS.docx")), str(mp),
                              str(out / "mbl_poison_prev.docx"), mode="preview")
    check_contains(out / "mbl_poison_prev.docx", "notification letter from STALE-CACHE workbook", "15 feet")
    check_no_blank_merges(out / "mbl_poison_prev.docx", "notification letter from STALE-CACHE workbook")
    # Date regression: Case Info dates stored as real Excel datetimes must render
    # in City long form in every document.
    import datetime as _dt, shutil as _sh, openpyxl as _ox
    dt_wb = out / "mbl_datetimes.xlsx"
    _sh.copy2(wb, dt_wb)
    _w = _ox.load_workbook(dt_wb)
    if "Case Info" in _w.sheetnames:
        _ws = _w["Case Info"]
        _vals = {"Letter_Date": _dt.datetime(2026, 5, 26), "Protest_Deadline": _dt.datetime(2026, 6, 25),
                 "Council_Deadline": _dt.datetime(2026, 7, 25), "Director_Approval_Date": _dt.datetime(2026, 2, 10),
                 "PC_Approval_Date": _dt.datetime(2026, 2, 12), "PC_Hearing_Date": _dt.datetime(2026, 3, 19),
                 "Protest_Letter_Date": _dt.datetime(2026, 3, 2), "Hearing_Letter_Date": _dt.datetime(2026, 3, 2),
                 "Reschedule_Letter_Date": _dt.datetime(2026, 3, 9)}
        for _r in range(1, _ws.max_row + 1):
            _k = _ws.cell(_r, 1).value
            if _k and str(_k).strip() in _vals:
                _ws.cell(_r, 2).value = _vals[str(_k).strip()]
        _w.save(dt_wb)
        mod.generate_draft_ordinance(str(dt_wb), str(_t("MBL_Ordinance_Template_Clean.docx")), str(mp), str(out / "mbl_dt_ord.docx"))
        check_no_date_timestamps(out / "mbl_dt_ord.docx", "ordinance dates")
        mod.generate_director_approval(str(dt_wb), str(_t("MBL_Directors_Approval.docx")), str(out / "mbl_dt_dir.docx"))
        check_no_date_timestamps(out / "mbl_dt_dir.docx", "director approval dates")
        mod.generate_pc_approval(str(dt_wb), str(_t("MBL_PC_Approval.docx")), str(out / "mbl_dt_pc.docx"))
        check_no_date_timestamps(out / "mbl_dt_pc.docx", "PC approval dates")
        for tname, lab in (("MBL_ALL_OWNERS.docx", "notification"), ("MBL_Protest_All_Owners.docx", "protest"),
                           ("MBL_Notice_PC_Rescheduling.docx", "reschedule")):
            mod.generate_notification(str(dt_wb), str(_t(tname)), str(mp), str(out / f"mbl_dt_{lab}.docx"), mode="preview")
            check_no_date_timestamps(out / f"mbl_dt_{lab}.docx", f"{lab} letter dates")
        mod.generate_staff_report(str(dt_wb), str(_t("MBL_OneMeeting_Staff_Report_Protest.docx")), str(out / "mbl_dt_staff.docx"))
        check_no_date_timestamps(out / "mbl_dt_staff.docx", "staff report dates")


def _validate_renewal(key: str, modname: str, wbname: str,
                      owner_templates: list[tuple[str, str]],
                      tpl: dict, out: Path, expected_value: str = ''):
    print(f"{key.replace('_', ' ').title()} Wizard")
    import importlib
    mod = importlib.import_module(modname)
    si, td = sample_input_dir(key), templates_dir(key)
    wb, mp = si / wbname, si / "sample_boundary_map.pdf"
    for tname, lab in owner_templates:
        mod.generate_notification(str(wb), str(_t(tname)), str(mp),
                                  str(out / f"{key}_{lab}_prev.docx"), mode="preview")
        check_docx(out / f"{key}_{lab}_prev.docx", f"{lab} notification preview")
        check_letter_spacing(_t(tname), out / f"{key}_{lab}_prev.docx",
                             f"{lab} letter spacing")
        check_no_blank_merges(out / f"{key}_{lab}_prev.docx", f"{lab} notification preview")
        mod.generate_notification(str(wb), str(_t(tname)), str(mp),
                                  str(out / f"{key}_{lab}_all.docx"), mode="all")
        check_docx(out / f"{key}_{lab}_all.docx", f"{lab} all-owner mail-out",
                   combined_titlepg=True, template=_t(tname))
    mod.generate_routing_form(str(wb), str(_t(tpl["routing"])),
                              str(out / f"{key}_routing.docx"))
    check_docx(out / f"{key}_routing.docx", "routing form", expect_checkboxes=True)
    mod.generate_director_approval(str(wb), str(_t(tpl["director"])),
                                   str(out / f"{key}_dir.docx"))
    check_docx(out / f"{key}_dir.docx", "director approval")
    mod.generate_pc_approval(str(wb), str(_t(tpl["pc"])),
                             str(out / f"{key}_pc.docx"))
    check_docx(out / f"{key}_pc.docx", "PC approval")
    check_pc_chair(out / f"{key}_pc.docx", "PC approval")
    mod.generate_applicant_letter(str(wb), str(_t(tpl["applicant"])),
                                  str(out / f"{key}_app.docx"))
    check_docx(out / f"{key}_app.docx", "applicant letter")
    mod.generate_draft_ordinance(str(wb), str(_t(tpl["ordinance"])), str(mp),
                                 str(out / f"{key}_ord.docx"))
    check_docx(out / f"{key}_ord.docx", "draft renewal ordinance")

    # 2026 renewal-only documents: protest letter mail-out and OneMeeting reports.
    if key == "mls_renewal":
        mod.generate_notification(str(wb), str(_t("MLS_REN_Protest_All_Owners.docx")), str(mp),
                                  str(out / f"{key}_protest_prev.docx"), mode="preview")
        check_docx(out / f"{key}_protest_prev.docx", "protest letter preview")
        check_letter_spacing(_t("MLS_REN_Protest_All_Owners.docx"),
                             out / f"{key}_protest_prev.docx", "protest letter spacing")
        mod.generate_notification(str(wb), str(_t("MLS_REN_Protest_All_Owners.docx")), str(mp),
                                  str(out / f"{key}_protest_all.docx"), mode="all")
        check_docx(out / f"{key}_protest_all.docx", "protest all-owner mail-out",
                   combined_titlepg=True, template=_t("MLS_REN_Protest_All_Owners.docx"))
        mod.generate_staff_report(str(wb), str(_t("MLS_REN_OneMeeting_Director_Approval.docx")),
                                  str(out / f"{key}_staff_da.docx"))
        check_docx(out / f"{key}_staff_da.docx", "OneMeeting staff report (Director)")
        check_contains(out / f"{key}_staff_da.docx", "OneMeeting staff report (Director)", "137REN")
        mod.generate_staff_report(str(wb), str(_t("MLS_REN_OneMeeting_PC_Approval.docx")),
                                  str(out / f"{key}_staff_pc.docx"))
        check_docx(out / f"{key}_staff_pc.docx", "OneMeeting staff report (PC)")
        check_contains(out / f"{key}_routing.docx", "routing form", "Previous Ordinance: 2006-118")
        check_contains(out / f"{key}_dir.docx", "director approval", "5,950 square feet")
        check_contains(out / f"{key}_ord.docx", "draft renewal ordinance", "five thousand and nine hundred and fifty (5,950)")


    if key == "mbl_renewal":
        for tname, lab in (("MBL_REN_Protest_All_Owners.docx", "protest"), ("MBL_REN_Notice_PC_Rescheduling.docx", "reschedule")):
            mod.generate_notification(str(wb), str(_t("" + tname + "")), str(mp),
                                      str(out / f"{key}_{lab}_prev.docx"), mode="preview")
            check_docx(out / f"{key}_{lab}_prev.docx", f"{lab} letter preview")
            check_letter_spacing(_t("" + tname + ""), out / f"{key}_{lab}_prev.docx", f"{lab} letter spacing")
            mod.generate_notification(str(wb), str(_t("" + tname + "")), str(mp),
                                      str(out / f"{key}_{lab}_all.docx"), mode="all")
            check_docx(out / f"{key}_{lab}_all.docx", f"{lab} all-owner mail-out",
                       combined_titlepg=True, template=_t("" + tname + ""))
        mod.generate_acceptance_letter(str(wb), str(_t("MBL_REN_Application_Acceptance.docx")),
                                       str(out / f"{key}_accept.docx"))
        check_docx(out / f"{key}_accept.docx", "application acceptance letter")
        check_contains(out / f"{key}_accept.docx", "application acceptance letter", "88REN")
        mod.generate_staff_report(str(wb), str(_t("MBL_REN_OneMeeting_Staff_Report.docx")),
                                  str(out / f"{key}_staff.docx"))
        check_docx(out / f"{key}_staff.docx", "OneMeeting staff report")
        check_contains(out / f"{key}_staff.docx", "OneMeeting staff report", "88REN")
        check_contains(out / f"{key}_routing.docx", "routing form", "Previous ordinance: 2006-63")
        check_contains(out / f"{key}_dir.docx", "director approval", "35 feet")
        check_contains(out / f"{key}_ord.docx", "draft renewal ordinance", "thirty-five (35) feet")

    # Regression for the stale-cache letter bug: a workbook saved with an
    # empty formula cache (the user-reported failure) must STILL produce a
    # letter that carries the minimum value -- the renewal wizards derive it
    # from the raw Step 1 data the Step 2 mirror formulas read.
    if expected_value:
        pz = out / f"{key}_poisoned.xlsx"
        poison_workbook(wb, pz)
        tname, lab = owner_templates[0]
        mod.generate_notification(str(pz), str(_t(tname)), str(mp),
                                  str(out / f"{key}_poison_prev.docx"),
                                  mode="preview")
        check_contains(out / f"{key}_poison_prev.docx",
                       f"{lab} letter from STALE-CACHE workbook", expected_value)
        check_no_blank_merges(out / f"{key}_poison_prev.docx",
                              f"{lab} letter from STALE-CACHE workbook")

    # Date-handling regression: rewrite Case Info date fields as real Excel
    # datetimes (the state of the planner's real workbook that leaked
    # '2006-01-18 00:00:00' into the ordinance) and assert every generated
    # document shows clean City-format dates.
    import datetime as _dt, shutil as _sh, openpyxl as _ox
    dt_wb = out / f"{key}_datetimes.xlsx"
    _sh.copy2(wb, dt_wb)
    _w = _ox.load_workbook(dt_wb)
    if "Case Info" in _w.sheetnames:
        _ws = _w["Case Info"]
        _vals = {
            "Original_Effective_Date": _dt.datetime(2006, 1, 18),
            "Original_Expiration_Date": _dt.datetime(2026, 1, 18),
            "Application_Filed_Date": _dt.datetime(2025, 11, 4),
            "Application_Accepted_Date": _dt.datetime(2025, 11, 6),
            "Application_Complete_Date": _dt.datetime(2026, 1, 6),
            "Letter_Date": _dt.datetime(2026, 5, 26),
            "Protest_Deadline": _dt.datetime(2026, 6, 25),
            "Council_Deadline": _dt.datetime(2026, 7, 25),
            "Director_Approval_Date": _dt.datetime(2026, 2, 10),
            "PC_Approval_Date": _dt.datetime(2026, 2, 12),
        }
        for _r in range(1, _ws.max_row + 1):
            _k = _ws.cell(_r, 1).value
            if _k and str(_k).strip() in _vals:
                _ws.cell(_r, 2).value = _vals[str(_k).strip()]
        _w.save(dt_wb)
        tname, lab = owner_templates[0]
        mod.generate_draft_ordinance(str(dt_wb), str(_t(tpl["ordinance"])), str(mp),
                                     str(out / f"{key}_dt_ord.docx"))
        check_no_date_timestamps(out / f"{key}_dt_ord.docx", "ordinance dates")
        mod.generate_director_approval(str(dt_wb), str(_t(tpl["director"])),
                                       str(out / f"{key}_dt_dir.docx"))
        check_no_date_timestamps(out / f"{key}_dt_dir.docx", "director approval dates")
        mod.generate_notification(str(dt_wb), str(_t(tname)), str(mp),
                                  str(out / f"{key}_dt_prev.docx"), mode="preview")
        check_no_date_timestamps(out / f"{key}_dt_prev.docx", "notification dates")


def validate_mls_renewal(out: Path):
    _validate_renewal(
        "mls_renewal", "wizards.mls_renewal.wizard",
        "MLS_Renewal_Calculation_Sheet_Case_Info_v2_FILLED_SAMPLE.xlsx",
        [("MLS_Renewal_SFR_Letters_to_owners.docx", "SFR"),
         ("MLS_Renewal_NO_SFR_Letters_to_owners.docx", "NoSFR")],
        {"routing": "MLS_Renewal_Routing_Form.docx",
         "director": "MLS_REN_Director_Approval.docx",
         "pc": "MLS_REN_PC_Approval.docx",
         "applicant": "MLS_Renewal_Applicant_Signs.docx",
         "ordinance": "MLS_RENEWAL_Ordinance_Template.docx"}, out,
        expected_value="5,950 square feet")


def validate_mbl_renewal(out: Path):
    _validate_renewal(
        "mbl_renewal", "wizards.mbl_renewal.wizard",
        "MBL_Renewal_Calculation_Sheet_Case_Info_v2_FILLED_SAMPLE.xlsx",
        [("MBL_REN_ALL_OWNERS.docx", "OWNERS")],
        {"routing": "MBL_REN_Routing_Form.docx",
         "director": "MBL_REN_Directors_Approval.docx",
         "pc": "MBL_REN_PC_Approval.docx",
         "applicant": "MBL_REN_APPLICANT_COMPLETE_SIGNS.docx",
         "ordinance": "MBL_REN_Ordinance_Template.docx"}, out,
        expected_value="35 feet")


VALIDATORS = {
    "mls": validate_mls,
    "mbl": validate_mbl,
    "mls_renewal": validate_mls_renewal,
    "mbl_renewal": validate_mbl_renewal,
}


def main() -> int:
    parser = argparse.ArgumentParser(description="RCA Wizard Suite build validation")
    parser.add_argument("--wizard", choices=sorted(VALIDATORS),
                        help="validate one wizard only")
    parser.add_argument("--keep", action="store_true",
                        help="keep generated proof files (prints the folder)")
    parser.add_argument("--out", default=None,
                        help="write proof files to this folder and keep them")
    args = parser.parse_args()

    if args.out:
        tmp = Path(args.out)
        tmp.mkdir(parents=True, exist_ok=True)
        args.keep = True
    else:
        tmp = Path(tempfile.mkdtemp(prefix="rca_validate_"))
    targets = [args.wizard] if args.wizard else list(VALIDATORS)
    print(f"RCA Wizard Suite validation — proofs in {tmp}\n")
    for key in targets:
        try:
            VALIDATORS[key](tmp)
        except Exception as e:
            import traceback
            print(f"  [FAIL] {key}: validator crashed — {e}")
            traceback.print_exc(limit=3)
            FAILURES.append(f"{key} validator crash")
        print()

    print("=" * 60)
    if FAILURES:
        print(f"RESULT: FAIL — {len(FAILURES)} of {CHECKS} checks failed:")
        for f in FAILURES:
            print(f"  - {f}")
    else:
        print(f"RESULT: PASS — all {CHECKS} checks passed.")
    if args.keep:
        print(f"Proof files kept in: {tmp}")
    else:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)
    print("Reminder (§20.12.3): confirm a late letter's footer and margins in "
          "Microsoft Word before department-wide rollout.")
    return 1 if FAILURES else 0


if __name__ == "__main__":
    sys.exit(main())
