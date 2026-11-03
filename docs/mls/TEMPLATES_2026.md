# MLS Wizard — 2026 City templates

The MLS (new application) wizard v15.0 fills the 2026 set of City templates
supplied by the Planning and Development Department in September 2026. The
previous token-based templates were archived in `templates/mls/_previous/`
(date-stamped) and are no longer used. The calculation logic (Step 2 reads,
70 % rule, support and SFR shares, planner-entered lot count) is unchanged;
only the template filling moved to `wizards/mls/templates_2026.py`, which uses
the shared engine in `core/docx_merge.py`.

## Template roles

| Role (manifest key) | File in `templates/mls/` | City source file |
|---|---|---|
| template | MLS_Owner_Notification_Letter.docx | MLS NOTIFICATION LETTERS_ALL OWNERS_LGL DESC (NEW TEMPLATE_2026).docx |
| protest_template *(new)* | MLS_Protest_All_Owners.docx | MLS PROTEST_ALL_OWNERS (NEW TEMPLATE) 2026.docx |
| hearing_template *(new)* | MLS_Notice_Public_Hearing_Less_51.docx | MLS NOTICE OF PUBLIC HEARING LESS 51% (NEW TEMPLATE) 2026.docx |
| applicant_template | MLS_Applicant_Complete_Signs.docx | APPLICANT_COMPLETE SIGNS (NEW TEMPLATE) 2026.docx |
| director_template | MLS_Director_Approval.docx | MLS Director_Approval.docx |
| pc_template | PC_Approval_NEW.docx | unchanged (no 2026 PC form was issued) |
| routing_template | MLS_Routing Form.docx | MLS_Routing Form.docx |
| staff_report_template *(new)* | MLS_OneMeeting_Staff_Report.docx | OneMeeting_MLS.docx |
| ordinance_template | MLS_Ordinance_Template.docx | MLS_Ordinance Template DA.docx |

The City's sample packet is kept as `docs/mls/MLS_RCA_Packet_SAMPLE_2026.pdf`
(routing form, Director approval, boundary map, tax map).

## Calculation workbook

`sample_input/mls/MLS_Calculation_Sheet_Case_Info_v13.xlsx` is the City's
2026 `NEW_Min_LotSize_Block_Calculation_Sheet.xlsx` with a **Case Info** tab
added (planner inputs, dropdowns, auto-read Step 2 values). Step 1 and Step 2
are untouched and have the same layout as before, so an in-progress v12
workbook still works; add the new Case Info rows to it when needed.
`..._v13_FILLED_SAMPLE.xlsx` is the demo case (No. 933, 5300-5600 Milart Street).

New Case Info fields (all optional, placeholders print when blank):

| Field | Used by |
|---|---|
| Land_Use_Deadline | owner letters: deadline for a corrected HCAD land use (defaults to Protest_Deadline) |
| Planner_Position, Council_Member, Applicant_Property_Address, Subdivision_Name | OneMeeting staff report |
| Protest_Count, Protest_Letter_Date, PC_Hearing_Date | protest letter, notice of public hearing |

## Placeholder wording the code matches

These templates are filled by matching the City's own placeholder wording
(manifest `fill: wording`). If the City rewrites one of these phrases the
matching line in `wizards/mls/templates_2026.py` needs the same change.

- Owner letters: `«OWNER_MAIL_To»` / `«Mailto»`, `«Mail_Addr_1»`, `«Mail_Addr_2»`,
  `«Mail_City»`, `«Mail_State»`, `«Mail_Zip»`, `«Legal_Dscr_1»`, `«group_dscr»`,
  `[Date]` (letter date paragraph), `X and X block of XX, X side, between XX and
  XX` / `x block of x Street, north and south sides,` / `between x Avenue and x
  Street`, `XXXX square feet` / `x square feet` / `XX`, `X%` (support),
  `[DATE]` in the land-use sentence and the protest sentence, `[Date]` in the
  hearing sentence, `Planner’s Name` / `Planner Name`, `832-393-XXXX`,
  `Name@houstontx.gov` / `Plannner@houstontx.gov`, `INSERT MAP`.
- Applicant letter: `Date`, `Name`, `Address`, `Houston, TX XXXX`, `XX block of
  XX Street, X and X sides,`, `between X Street and X Drive`, `Property Owner`
  (salutation), `xx square feet`, `two (2) signs`, `one sign at each corner`,
  `Planner Name at (832) 393-XXXX`, `Planner.Name@houstontx.gov`.
- Director approval: `XXX` (block number), the red area line, `XXXX square
  feet`, `Prepared by: Planner Name`, `Vonn Tran`, `Date`.
- Routing form (content controls, by position): topic, location, author,
  `DATE`, `#XXX`, `: XXXX` (lot size), council district `XXX`; checkboxes LOT
  SIZE, 60 % SFR Yes/No, SMLS Block Application, PC / Director's Approval.
- OneMeeting report: `XXX`, `2800 - 2900 block of Wheeler Street…`, `ENTER THE
  CAPTION…`, `DISTRICT D - EVANS-SHABAZZ`, `XXXX XXX Street, of the SUBDIVISION
  NAME`, the block/side/cross-street run sequence, `Your Name, Position, Phone
  Number`.
- Ordinance (content controls): `MLS Boundary Description` (tag), the
  lower-case area control, `51%`, `sixteen (16) out of twenty-five (25)`, `six
  thousand eight hundred and twenty (6,820)` (three times), `6,820 feet.`
  (filled as `N square feet.`), `XXXXXXXXX` (L.D. file, only when
  Legal_File_No is given), Exhibit A / B / C block controls.

## RCA packet (2026)

Order: routing form, Director or Planning Commission approval, boundary map,
tax map, *(optional: subdivision plat, Step 2 staff analysis, HCAD records)*.
Only the tax map must be uploaded; this matches the attachment list on the
2026 routing form and the City's sample packet.
