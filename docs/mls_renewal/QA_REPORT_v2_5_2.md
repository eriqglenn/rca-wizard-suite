# MLS Renewal Wizard v2.5.2 QA / Source-Brief Compliance Pass

## Purpose

This version is a QA and source-brief compliance pass for the MLS Renewal Wizard. It does not add a new major feature. It verifies the existing modules against the project build rules and includes the source brief inside the package as `docs/BUILD_RULES.md`.

## Source-brief rules applied

- Use official City templates as the formatting source.
- Do not recreate City forms or globally normalize fonts.
- Preserve Step 1 and Step 2 workbook structure.
- Generate Word documents for editable forms and letters.
- Generate PDF only for RCA packet assembly, staff-analysis export, and proofing.
- Keep Draft Renewal Ordinance as standalone Word output.
- Include Original Ordinance PDF inside the Renewal RCA packet.
- Export only meaningful Step 2 staff-analysis content.
- Include visual previews for changed/important outputs.

## QA run performed

Generated using the filled sample renewal workbook and sample attachments:

- SFR Renewal Notification Preview
- No-SFR Renewal Notification Preview
- Applicant Sign Letter
- Director Approval
- Planning Commission Approval
- Renewal Routing Form
- Draft Renewal Ordinance
- Step 2 Staff Analysis PDF
- MLS Renewal RCA Packet PDF

## Issue caught during QA

The routing form still had stale red placeholder content controls layered on top of the filled visible text. This was corrected in v2.5.2 by deleting the stale routing-form content controls after the visible template paragraphs are filled. This is limited to routing form cleanup and does not rebuild the form.

## Visual QA notes

- SFR letter: City logo/letterhead preserved; Spanish notice remains on page 1; map follows on page 2.
- No-SFR letter: City logo/letterhead preserved; Spanish notice remains on page 1; map follows on page 2.
- Applicant letter: City letterhead preserved.
- Director Approval: renewal-specific criteria and bold findings preserved.
- PC Approval: renewal-specific criteria preserved.
- Routing Form: SMLS Block Renewal and Original Ordinance checked; stale red placeholders removed.
- RCA Packet: order is Routing Form, Director/PC Approval, Boundary Map, Staff Analysis, Subdivision Plat, Tax Map, HCAD Records, Original Ordinance.
- Draft Renewal Ordinance: generated as standalone Word document; not inserted into RCA packet.

## Known sample-data note

The sample workbook still contains placeholder/sample values, including Application No. `0`. That is source sample data, not a packet-builder error.

