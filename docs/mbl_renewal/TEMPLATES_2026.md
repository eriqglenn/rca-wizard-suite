# MBL Renewal Wizard — 2026 City templates

The MBL Renewal wizard (v2.0) fills the 2026 set of City templates supplied
by the Planning and Development Department in September 2026. The previous
templates were archived in `templates/mbl_renewal/_previous/` (date-stamped).
The wizard was rewritten on the shared wording-based engine in
`core/docx_merge.py` (ordinance paragraph and exhibit helpers are shared with
the MLS Renewal wizard).

## Template roles

| Role (manifest key) | File in `templates/mbl_renewal/` | City source file |
|---|---|---|
| owner_template | MBL_REN_ALL_OWNERS.docx | MBL REN_ALL OWNERS (NEW TEMPLATE 2026).docx |
| protest_template *(new)* | MBL_REN_Protest_All_Owners.docx | MBL_REN PROTEST_PC ALL OWNERE (NEW TEMPLATE 2026).docx |
| reschedule_template *(new)* | MBL_REN_Notice_PC_Rescheduling.docx | MBL_NOTICE OF PC_RESCHEDULING (NEW TEMPLATE 2026).docx |
| acceptance_template *(new)* | MBL_REN_Application_Acceptance.docx | MBL_APPLICATION ACCEPTANCE (NEW TEMPLATE) 2026.docx |
| applicant_template | MBL_REN_APPLICANT_COMPLETE_SIGNS.docx | MBL_REN APPLICANT COMPLETE_SIGNS (NEW TEMPLATE) 2026.docx |
| director_template | MBL_REN_Directors_Approval.docx | MBL_REN Directors Approval Template.doc (converted to .docx by Word; source kept in `docs/mbl_renewal/`) |
| pc_template | MBL_REN_PC_Approval.docx | unchanged (no 2026 PC form was issued) |
| routing_template | MBL_REN_Routing_Form.docx | MBL_Routing Form.docx |
| staff_report_template *(new)* | MBL_REN_OneMeeting_Staff_Report.docx | OneMeeting_MBL_RENEWAL.docx |
| ordinance_template | MBL_REN_Ordinance_Template.docx | MBL_REN_Ordinance Template.docx |

## Calculation workbook (2026 layout)

`sample_input/mbl_renewal/MBL_Renewal_Calculation_Sheet_Case_Info_v2.xlsx` is
the City's `RENEW_Min_Building_Line_Calculation_Sheet_NEW.xlsx` with a **Case
Info** tab added; `..._FILLED_SAMPLE.xlsx` is the demo case (88REN, 2400-2500
blocks of Rosedale Street). Step 2 cells the wizard reads:

| Cell / range | Meaning |
|---|---|
| B2 | Application No. (`88` → the wizard prints `88REN`) |
| D2 | Building Line to Renew (feet) |
| B3 / D3, B4, D4, B5 / D5, B6 / D6 | dates, street, block description, cross streets, sides |
| A9 / A13 / D9 / D10 | SFR lots / total lots / SFR share / qualification statement |
| A15:F97 | property table; C original line, D current line, F `Yes` when the line is unchanged or meets the line to renew |

The pre-2026 sheet (E10 most frequent line, H9 share) is still recognised, so
in-progress cases keep working. Counts used by the PC approval (lots at the
renewed line, lots varying by more than five feet) are computed from column D
in both layouts.

New Case Info fields: `Planner_Position`, `Council_Member`,
`Applicant_Property_Address`, `Subdivision_Name`, `Protest_Count`,
`Protest_Letter_Date`, `PC_Hearing_Date`, `Reschedule_Letter_Date`,
`Support_Percentage` (rescheduling notice only). Blank values keep the
template's visible placeholder.

## Placeholder wording the code matches

- Owner letters: `«Owner_Mail»`, `«Mail_Addr_»`, `«Mail_City»`, `«Mail_State»`,
  `«Mail_Zip»`, `DATE` (letter date paragraph; also the protest deadline and
  the new hearing date inside sentences), `XX block of XX Street, X and X
  sides,`, `between X and X`, `XX feet`, `XX%` (support), `Thursday, [Date]`,
  `Planner Name` / `contact Planner,`, `832-393-XXXX` / `832-393-XXX`,
  `PlannerName@houstontx.gov` / `Plannner@houstontx.gov`, `INSERT MAP`.
- Applicant letters (acceptance and sign letter): the same `«…»` fields
  filled with the applicant's address, `DATE`, `XXX block of XX Street, XX and
  XX sides,`, `between XX and XX`, `Property Owners Name`, the acceptance
  sentence run sequence (`Date`, `XX Blk of XX (X side) Between XX and XX, is
  MBL #XXX`, `Planner’s Name`), `XX feet`, `two (2) signs`, `one sign at each
  corner`, `Planner at (832) 393-XXXX`, `Planner Name@houstontx.gov`.
- Director approval: `XXREN`, the `XX block of XX Street, X side between X
  and X Street` run sequence, `2006-63`, `January 31, 2006`, `XX feet`, `Vonn
  Tran`, `Date`.
- PC approval (unchanged form): the finding lines starting `The application
  is for the`, `% of the lots are developed`, `protest was filed`, `The block is
  comprised of`, `The application meets this standard`, and `Date`.
- Routing form (content controls): topic, location, author, `DATE`, `Previous
  ordinance: XXXX`, `#XXX`, `: XX` (building line), `XX` (SFR %), ` XX`
  (petition → N/A), `XX / XX` (total / meeting); checkboxes BUILDING LINE,
  SMBL Block Renewal, PC / Director's Approval.
- OneMeeting report: `XXX`, `2800 - 2900 block of Wheeler Street…`, `ENTER THE
  CAPTION…`, `DISTRICT D - EVANS-SHABAZZ`, `XX Street of the (Subdivision
  Name)`, the block/side/cross-street run sequence, `(200X-XXX)`, `20__.`,
  `Your Name, Position, Phone Number`.
- Ordinance: the `MLS Boundary Description` caption control (same tag as the
  MLS form), the `WHEREAS` recitals (original ordinance, application dates,
  protest status, Director finding), `Section 4`, `Requested by`, Exhibits A–C.

## RCA packet

Routing form, Director or Planning Commission approval, boundary map, tax map,
*(optional: subdivision plat, Step 2 staff analysis, HCAD records)*, original
ordinance. Only the tax map and the original ordinance must be uploaded.
