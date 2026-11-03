# RCA Wizard Suite (MLS Wizard v30)

A Windows desktop application I built while working as a City Planner at the City of Houston Planning & Development Department. It turns a planner's calculation workbook into the complete legal document set for a Special Minimum Lot Size (MLS) or Special Minimum Building Line (MBL) case. That work used to take days of hand-editing Word templates. With the suite it takes minutes.

**About 34,000 lines of Python · 26 test modules · packaged as a standalone .exe (no Python install needed)**

## What it does

One program with a home screen and four wizards:

| Wizard | Case type |
|---|---|
| MLS Wizard | Special Minimum Lot Size, new block application |
| MBL Wizard | Special Minimum Building Line, new block application |
| MLS Renewal Wizard | Special Minimum Lot Size, block renewal |
| MBL Renewal Wizard | Special Minimum Building Line, block renewal |

From one Excel calculation sheet plus a few attachments, each wizard produces:

- **Owner notification letters** for every property owner, merged into one mail-out with the boundary map after each letter
- **Protest and public-hearing notices**, chosen automatically from the petition's support percentage
- **Director or Planning Commission approvals**, the routing form and the OneMeeting staff report
- **A draft ordinance** with Exhibits A, B and C
- **The final RCA packet PDF:** routing form, approval, boundary map, tax map, plat, staff analysis and county records, merged and condensed
- **Planning Commission materials:** draft report, speaking notes and land-use map

## Engineering highlights

- **Boundary maps from live GIS data** (`core/boundary_map*.py`). The wizard pulls parcel polygons and street centerlines from Harris County and City ArcGIS REST services. It fits labels inside irregular parcels, and renders a print-ready vector PDF plus a PNG with an interactive preview for adjusting framing and labels. Every map saves a provenance file recording data sources and retrieval dates.
- **Application import with OCR** (`core/application_import*.py`, `application_ocr.py`). It reads a scanned application PDF, matches it to HCAD parcels and fills in the calculation workbook.
- **Template engine** (`core/docx_merge.py`, `city_templates.py`). It does token replacement across Word runs, tables, headers and footers. It also keeps formatting through mail merges, and checks every template against a manifest before generating anything.
- **Office automation with fallbacks** (`core/office_convert.py`, `soffice.py`, `pdf_setup.py`). It converts to PDF through desktop Word/Excel (COM) or LibreOffice, and has a built-in diagnostic for locked-down work computers.
- **Browser automation** (`core/hcad_downloader.py`, `ilms_browser.py`, `plat_search.py`). It uses Playwright to download county property records and place permit holds in the City's permitting system. It never reads or stores credentials. The planner signs in, and the tool only drives a visible browser window.
- **Peer reviewer** (`core/peer_review*.py`). It re-reads the generated documents and flags inconsistencies (counts, dates, addresses) before a packet is routed.
- **Defensive by design.** Missing or ambiguous data stops generation with a plain-English message instead of producing a wrong document. Cached data is labeled with retrieval dates.

## Layout

```
app.py              launcher and home screen (Tkinter)
core/               shared engine: GIS maps, templates, Office/PDF, OCR, browser automation, review
wizards/<type>/     one package per wizard: UI flow, mail-out engine, PC report/speech builders
tools/              unit and integration tests, setup checks, template onboarding
docs/               developer notes per release, feature docs, user guide
packaging/          PyInstaller spec and build script for the standalone Windows package
```

## What is not in this repository

Some material is left out on purpose. Because of that, some tests that need sample workbooks, templates or the seal will fail on a fresh clone.

- **City document templates and the City seal:** City of Houston property.
- **Sample case workbooks and maps:** these contain real property and mailing-address records.
- **Internal City server addresses:** replaced with `*.internal.example` placeholders.
- **Executables and build output.**

## Stack

Python 3 · Tkinter · python-docx · openpyxl · PyMuPDF · pypdf · ReportLab · Shapely · Pillow · Playwright · pywin32 · PyInstaller
