# MBL Wizard — 2026 City templates (suite v1.15 / v1.10.4)

The MBL (new application) wizard v9.0 fills the City's 2026 Special Minimum
Building Line Block templates **by wording**: the City files are installed
unchanged in `templates\mbl` and the wizard replaces the City's own
placeholder text (`XXX block of XX Street`, `Planner Name`, `832-393-XXXX`,
`DATE`, content controls on the routing form) inside the existing runs.
Nothing in the templates has to be tagged with `{{tokens}}`.

The previous MBL templates are archived in `templates\mbl\_previous`
(`20260904_140000__*`).

## Files and roles

| Role key | File (templates\mbl) | City source | Used for |
|---|---|---|---|
| template | MBL_ALL_OWNERS.docx | MBL ALL OWNERS (NEW TEMPLATE) 2026.docx | Owner notification letter (30-day protest period) |
| hearing_template | MBL_Notice_Public_Hearing_Less_51.docx | MBL NOTICE OF PUBLIC HEARING LESS 51% (NEW TEMPLATE) 2026.docx | Notice of public hearing when support is under 51% |
| protest_template | MBL_Protest_All_Owners.docx | MBL PROTEST_PC_ALL_OWNERS (NEW TEMPLATE) 2026.docx | Protest / public hearing letter to all owners |
| reschedule_template | MBL_Notice_PC_Rescheduling.docx | MBL_NOTICE OF PC RESCHEDULING (NEW TEMPLATE) 2026.docx | New hearing date to all owners |
| applicant_template | MBL_APPLICANT_COMPLETE_SIGNS.docx | MBL APPLICANT_COMPLETE_SIGNS (NEW TEMPLATE) 2026.docx | Completion / sign-posting letter |
| director_template | MBL_Directors_Approval.docx | MBL _Directors Approval New Template (2026).doc (converted with Word; source in docs\mbl) | Director approval (support 51% or more, no protest) |
| pc_template | MBL_PC_Approval.docx | MBL_PC_Approval_Template.doc (converted with Word; source in docs\mbl) | Planning Commission approval with findings |
| routing_template | MBL_Routing_Form.docx | MBL_Routing Form New Template.docx | 2026 coversheet routing form |
| staff_report_template | MBL_OneMeeting_Staff_Report.docx | OneMeeting_MBL.docx | OneMeeting staff report, Director path |
| staff_report_less51_template | MBL_OneMeeting_Staff_Report_Less51.docx | OneMeeting_MBL_LESS 51%.docx | OneMeeting staff report, support under 51% |
| staff_report_protest_template | MBL_OneMeeting_Staff_Report_Protest.docx | OneMeeting_MBL_PROTEST.docx | OneMeeting staff report, protest filed |
| ordinance_template | MBL_Ordinance_Template_Clean.docx | (unchanged) | Draft ordinance, token engine |

**Draft ordinance.** Both copies of the City's `MBL_Ordinance Template.docx`
in the 2026 set are password-encrypted Office packages (Word and LibreOffice
cannot open them without the password), so the ordinance still runs on the
previous token template. Drop in the unlocked 2026 file and the ordinance can
be moved to the wording engine like the MBL Renewal ordinance.

## Workbook

`sample_input\mbl\MBL_Calculation_Sheet_Case_Info_v3.xlsx` is the City's
2026 `NEW_Min_Building_Line_Calculation_Sheet.xlsx` (Step 1 now carries an
`HCAD_NUM_1` column) with a **Case Info** tab added. Yellow cells are planner
inputs, green cells are dropdowns, blue cells are live formulas that read
Step 2. The filled sample (`..._FILLED_SAMPLE.xlsx`, application 1030, 1800
block of Haver Street) drives the validator.

Auto-read from Step 2: Application_Number (B2), Date_Received (B3),
Date_Complete (D3), Street_Name (B4), Block_Description (D4), Side_Of_Street
(B6/D6), Cross_Streets (B5/D5), Minimum_Building_Line (L16, the line at least
70% of developed lots meet), Support_Percentage (F8), SFR_Percentage (H11),
SFR_Count (A11), Total_Lot_Count (A15), Developed_Lot_Count (AC17),
Qualifying_Lot_Count (lots with a line of at least L16), Qualifies (E3).
When a workbook is saved with an empty formula cache the wizard recomputes all
of these from the typed Step 1 / Step 2 columns.

Planner inputs: Letter_Date, Protest_Deadline, Planner_Name / Position /
Phone / Email, Council_District, Council_Member, Council_Deadline,
Approval_Path, Applicant_* (name, address, property address),
Subdivision_Name, Sign_Count_*, Sign_Placement, Director_Name,
Director_Approval_Date, PC_Signer, PC_Approval_Date, Subdivision_Plat_Year,
Neighborhood_Age_Statement, Structure_Age_Statement, Preservation_Finding,
Protest_Count, Protest_Letter_Date, Hearing_Letter_Date, PC_Hearing_Date,
Reschedule_Letter_Date, Ordinance_Year, Legal_File_No.

## Placeholder map (what the wizard replaces)

Owner letters (all four): merge fields `«Owner_Mail»`, `«Mail_Addr_»`,
`«Mail_City» «Mail_State» «Mail_Zip»`; `DATE` paragraph (Letter_Date,
Hearing_Letter_Date, Protest_Letter_Date or Reschedule_Letter_Date by
letter); subject lines `XXX block of XX Street, XX and XX sides,` /
`between XX and XX`; `X feet` / `XX feet` (building line); `Month XX, 2026`
(Protest_Deadline); `X%` / `XX%` (Support_Percentage); `[Date]` /
`DATE at 2:30 p.m.` (PC_Hearing_Date, weekday for `Thursday,`);
`Planner Name` / `Your Name` / `832-393-XXXX` / e-mail placeholders;
`INSERT MAP` (boundary map).

Applicant letter: merge fields (applicant), `DATE`, subject lines,
`Property Owners Name`, `X feet`, `two (2)`, `one sign at each corner`,
`Planner`, `(832) 393-XXXX`, `Planner Name@houstontx.gov`.

Director approval: `Block No. XX`, `minimum building line of XXXX feet`,
`Vonn Tran`, `Date`.

PC approval: `Area No. 106`, `The application is for the ...`, `88% of the
proposed application area ...`, `Petition signed by owners of 53% ...`,
`A minimum building line of 15 feet exists on ten (10) properties ...`,
`The subdivision was platted in xxxx ...` (Neighborhood / Structure /
Preservation statements), `Ten (10) out of thirteen (13) developed
properties ...`, `Date`. Both signature blocks (Chair / Vice-Chair) stay.

Routing form (content controls): topic, location, author, Council deadline,
Building Line checkbox, `#` application number, `Established: XXXX`,
`SFR: XXX%`, `Petition Signed By: XX%`, `Total Lots / Lots That Meet
Building Line`, RCA Applies to (SMBL Block Application), attachments
(Director or PC approval, map, staff analysis, plat, tax map, tax records).

OneMeeting staff reports: `CHOOSE THE DISTRICT`, `Application No. XXX`,
item title location, item caption (ordinance caption + `DISTRICT X - NAME`),
background paragraphs (applicant property, subdivision, area, support,
protest sentence, hearing date), contact line. The wizard picks the variant
from Case Info: protest filed → protest report; support under 51% →
under-51% report; otherwise the Director report.
