# MLS Renewal Wizard — 2026 City templates

The MLS Renewal wizard (v3.0) fills the 2026 set of City templates supplied
by the Planning and Development Department in September 2026. The previous
templates were archived in `templates/mls_renewal/_previous/` (date-stamped)
and are no longer used.

## Template roles

| Role (manifest key) | File in `templates/mls_renewal/` | City source file |
|---|---|---|
| sfr_template | MLS_Renewal_SFR_Letters_to_owners.docx | MLS RENEWAL_SFR_LTR TO OWNERS MAIL_MERGE (NEW TEMPLATE) 2026.docx |
| no_sfr_template | MLS_Renewal_NO_SFR_Letters_to_owners.docx | MLS RENEWAL_NO SFR_LTR TO OWNERS MAIL_MERGE (NEW TEMPLATE) 2026.docx |
| protest_template *(new)* | MLS_REN_Protest_All_Owners.docx | MLS_REN_PROTEST_ALL OWNERS.docx |
| applicant_template | MLS_Renewal_Applicant_Signs.docx | MLS RENEWAL_APPLICANT_SIGNS.docx |
| director_template | MLS_REN_Director_Approval.docx | MLS_REN_Director Approval.docx |
| pc_template | MLS_REN_PC_Approval.docx | unchanged (no 2026 PC form was issued) |
| routing_template | MLS_Renewal_Routing_Form.docx | MLS_Routing Form.docx |
| staff_report_director_template *(new)* | MLS_REN_OneMeeting_Director_Approval.docx | OneMeeting_MLS_RENEWAL.docx |
| staff_report_pc_template *(new)* | MLS_REN_OneMeeting_PC_Approval.docx | PC_ ONE MEETING FOR MLS RENEWAL.docx |
| ordinance_template | MLS_RENEWAL_Ordinance_Template.docx | MLS_RENEWAL_Ordinance Template.docx |

The City's sample packet is kept as `docs/mls_renewal/MLS_REN_RCA_Packet_SAMPLE_2026.pdf`.

## Calculation workbook (2026 layout)

`sample_input/mls_renewal/MLS_Renewal_Calculation_Sheet_Case_Info_v2.xlsx` is
the City's `RENEW_Min_Lot_Size_Calculation_Sheet_NEW.xlsx` with a **Case Info**
tab added (planner inputs, dropdowns, auto-read Step 2 values). Step 1 and
Step 2 are untouched. `..._v2_FILLED_SAMPLE.xlsx` is the same workbook filled
for the demo case (137REN, 4400 block of Leeland Street).

Step 2 cells the wizard reads (`Step 2 - Main Calculation Sheet`):

| Cell / range | Meaning |
|---|---|
| B2 | Application No. (`137` → the wizard prints `137REN` on titles) |
| D2 | Lot Size to Renew (square feet) = the minimum lot size |
| B3 / D3 | Date Received / Date Complete |
| B4, D4 | Street name, block description (`4400 block of Leeland Street`) |
| B5, D5 | Cross streets |
| B6, D6 | Sides of street |
| A9 / A13 / D9 | SFR lots / total lots / SFR share |
| D10 | Qualification statement |
| A15:F97 | Property table; column F = `Yes` when the lot is unchanged or meets the lot size |

The 2026 sheet no longer contains the corner-lot / prevailing-lot-size (75 %
within 10 %) analysis, so the approvals and ordinance use the lot-size-to-renew
figure and the count of lots meeting it (`Lots_Meeting_Minimum`).

### Case Info fields

| Field | Used by |
|---|---|
| Letter_Date, Protest_Deadline, Planner_Name, Planner_Phone, Planner_Email | owner letters, applicant letter, routing form |
| Planner_Position, Council_District, Council_Member | routing form, OneMeeting staff reports |
| Council_Deadline | routing form ("Council should consider … no later than") |
| Approval_Path, Owner_Letter_Type | which approval / letter / staff-report template is used |
| Applicant_* , Sign_Count_*, Sign_Placement | applicant sign letter |
| Applicant_Property_Address, Subdivision_Name | staff reports ("the property owner at … of the … Subdivision") |
| Director_Name, Director_Approval_Date, PC_Signer, PC_Approval_Date | approvals |
| Protest_Count, Protest_Letter_Date, PC_Hearing_Date | protest letter, PC approval, PC staff report, ordinance recital |
| Original_Ordinance_No, Original_Effective_Date, Original_Expiration_Date | Director approval, routing form, ordinance, staff reports |
| Application_Filed/Accepted/Complete_Date, Ordinance_Year, Legal_File_No | draft ordinance |

Blank dates print the word `Date` / keep the template's red placeholder so the
planner notices; nothing is silently dropped.

## Placeholder wording the code matches

These wizards fill by matching the template's own wording (manifest `fill:
wording`). If the City rewrites one of these phrases, the matching line in
`wizards/mls_renewal/wizard.py` needs the same change.

- Owner letters: `«Owner_Mail»`, `«Mail_Addr_»`, `«Mail_City»`, `«Mail_State»`,
  `«Mail_Zip»`, `[DATE]` (letter date paragraph and protest deadline),
  `XXX block of XX Street, XX and XX sides,`, `between XX and XX Streets`,
  `XXXX` / `X, XXX` square feet, `Planner’s Name`, `832-393-XXXX`,
  `Name@houstontx.gov`, `INSERT MAP`.
- Protest letter: as above plus `DATE`, `Thursday, [Date] at 2:30 p.m.`,
  `Planner Name`, `Plannner@houstontx.gov`.
- Applicant letter: `DATE`, `APPLICANT NAME`, `ADDRESS`, `Houston, TX 77XXX`,
  `XX Block of XX, X side,`, `between XX and XX.`, `XX square feet`,
  `two (2) signs`, `one sign at each corner`, `Planner at (832) 393-xxxx`,
  `Planner Name@houstontx.gov`.
- Director approval: `XXREN`, `XX block of XX Street (X side); between X and X
  Streets`, `XXXX square feet`, `2005-729`, `June 14, 2005`, `Vonn Tran`, `Date`.
- Routing form (content controls): `Special Minimum Lot Size/ Block Renewal
  Application No. X`, `The X block of X Street, X and X sides, between X Street
  and X Street`, `Planner Name, Phone Number`, `[DATE]`, `Previous Ordinance:
  XXXX`, `XXX`, `: XXXX`, ` X` (council district). Checkbox controls are set
  by position (LOT SIZE, SMLS Block Renewal, Planning Commission / Director's
  Approval).
- OneMeeting reports: the red sample values (`XXX`, `2800 - 2900 block of
  Wheeler Street…`, `DISTRICT D - EVANS-SHABAZZ`, `(200X-XXX)`, `Your Name,
  Position, Phone Number`, `Month X, 202X`, `CM NAME`, …) filled run by run in
  document order.
- Ordinance: the `MLS Boundary Description` content control (caption), the
  `WHEREAS` recitals naming the original ordinance, application dates, protest
  status and the lot-size finding, `Section 4`, `Requested by`, Exhibits A–C.

## RCA packet (2026)

Order: routing form, Director or Planning Commission approval, boundary map,
tax map, *(optional: subdivision plat, Step 2 staff analysis, HCAD records)*,
original ordinance. Only the tax map and original ordinance must be uploaded;
this matches the attachment list on the 2026 routing form and the City's
sample packet.
