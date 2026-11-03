"""RCA Wizard Suite — template onboarding inspector (master brief §20.0).

Inspects every City template in ``Wizard_Files/templates/<wizard>/`` and
writes one onboarding report per template into ``Wizard_Files/docs/onboarding``.
Each report records the structural facts the merge code depends on:

  - every literal {{token}} in the body, headers, and footers
  - section count, per-section margins, break types, titlePg flags
  - header/footer references (including first-page footers)
  - checkbox mechanism in use (content control / legacy form field / glyph)

Why this matters: when the City revises a template, re-run this tool and diff
the new report against the saved one.  A changed token list, dropped footer
reference, or different checkbox type tells you exactly what merge code (or
field map) needs attention *before* a bad document reaches a planner.

Run from Wizard_Files:

    python tools/onboard_templates.py            # all wizards
    python tools/onboard_templates.py --wizard mls
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from datetime import datetime
from pathlib import Path

WIZARD_FILES = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(WIZARD_FILES))

from core.paths import TEMPLATES_ROOT, WIZARD_FILES_DIR  # noqa: E402

REPORT_DIR = WIZARD_FILES_DIR / "docs" / "onboarding"


def inspect_template(docx_path: Path) -> str:
    z = zipfile.ZipFile(docx_path)
    parts = [n for n in z.namelist() if n.startswith("word/")]
    body_xml = z.read("word/document.xml").decode("utf-8", errors="replace")

    # tokens in body + every header/footer part
    token_sources = {"document.xml": body_xml}
    for part in parts:
        if re.match(r"word/(header|footer)\d*\.xml$", part):
            token_sources[part.split("/")[-1]] = z.read(part).decode(
                "utf-8", errors="replace")

    lines = []
    lines.append(f"TEMPLATE: {docx_path.name}")
    lines.append(f"INSPECTED: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append("")

    lines.append("TOKENS:")
    any_token = False
    for src, xml in token_sources.items():
        tokens = sorted(set(re.findall(r"\{\{[^}]+\}\}", xml)))
        if tokens:
            any_token = True
            lines.append(f"  in {src}:")
            for t in tokens:
                lines.append(f"    {t}")
    if not any_token:
        lines.append("  (none — this template is filled by label/paragraph "
                     "matching, not {{token}} replacement)")
    lines.append("")

    # split-run token risk: does '{{' appear without a matching full token in
    # any single w:t run? (heuristic: count '{{' in raw xml vs in joined text)
    raw_open = body_xml.count("{{")
    joined = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", body_xml))
    full_in_runs = len(re.findall(r"\{\{[^}]+\}\}", body_xml))
    full_in_text = len(re.findall(r"\{\{[^}]+\}\}", joined))
    if full_in_text > full_in_runs:
        lines.append(f"SPLIT-RUN RISK: YES — {full_in_text - full_in_runs} "
                     "token(s) only visible after joining runs (see brief §4; "
                     "use run-reassembly or fix the template).")
    else:
        lines.append("SPLIT-RUN RISK: none detected (each token sits inside a "
                     "single run).")
    lines.append("")

    sects = re.findall(r"<w:sectPr.*?</w:sectPr>", body_xml, re.S)
    lines.append(f"SECTIONS: {len(sects)}")
    for i, s in enumerate(sects):
        def attr(pattern, default="?"):
            m = re.search(pattern, s)
            return m.group(1) if m else default
        btype = attr(r'<w:type w:val="(\w+)"', "nextPage(default)")
        left, right = attr(r'w:left="(\d+)"'), attr(r'w:right="(\d+)"')
        top, bottom = attr(r'w:top="(\d+)"'), attr(r'w:bottom="(\d+)"')
        frefs = re.findall(r'footerReference[^>]*w:type="(\w+)"', s)
        hrefs = re.findall(r'headerReference[^>]*w:type="(\w+)"', s)
        titlepg = "yes" if "w:titlePg" in s else "no"
        lines.append(f"  sect{i}: break={btype} margins(L/R/T/B)="
                     f"{left}/{right}/{top}/{bottom} titlePg={titlepg} "
                     f"footers={frefs or '-'} headers={hrefs or '-'}")
    lines.append("")

    cc = body_xml.count("w14:checkbox")
    legacy = body_xml.count("FORMCHECKBOX")
    glyphs = body_xml.count("\u2610") + body_xml.count("\u2612")
    lines.append("CHECKBOXES:")
    lines.append(f"  content-control (w14:checkbox): {cc}")
    lines.append(f"  legacy form field (FORMCHECKBOX): {legacy}")
    lines.append(f"  plain glyph characters (\u2610/\u2612): {glyphs}")
    if cc:
        lines.append("  -> patch via OOXML: set BOTH w14:checked AND the glyph "
                     "(brief §8).")
    elif glyphs:
        lines.append("  -> glyph replacement applies.")
    lines.append("")

    lines.append(f"PARTS: {', '.join(parts)}")
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description="Template onboarding reports")
    parser.add_argument("--wizard", help="one wizard key (default: all)")
    args = parser.parse_args()

    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    wizards = ([args.wizard] if args.wizard
               else sorted(p.name for p in TEMPLATES_ROOT.iterdir() if p.is_dir()))
    written = 0
    for wkey in wizards:
        folder = TEMPLATES_ROOT / wkey
        for tpl in sorted(folder.glob("*.docx")):
            try:
                report = inspect_template(tpl)
            except Exception as e:
                report = f"TEMPLATE: {tpl.name}\nERROR: could not inspect — {e}\n"
            out = REPORT_DIR / f"{wkey}__{tpl.stem}.txt"
            out.write_text(report, encoding="utf-8")
            written += 1
            print(f"wrote {out.name}")
    print(f"\n{written} onboarding report(s) in {REPORT_DIR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
