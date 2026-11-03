# Source Brief: How to Build City Planning RCA Wizards

## 1. Core Philosophy

The wizard is not supposed to redesign City forms. It is a controlled document-generation tool that fills existing City templates with values from a calculation workbook and user-provided attachments.

The central rule is:

**Use the official templates exactly as provided. Do not recreate them, redesign them, or normalize their fonts globally. Only replace specific case-related placeholder text while preserving the original template formatting.**

The wizard should behave like a sophisticated mail merge plus packet assembler.

It should:

1. Read owner/property data from the calculation sheet.
2. Read case-level data from a `Case Info` tab.
3. Pull calculation results from Step 2.
4. Fill City Word templates without changing formatting.
5. Insert selected maps where required.
6. Generate Word documents for editable forms/letters.
7. Generate PDFs only where packet assembly or proofing requires it.
8. Assemble RCA packet PDFs from generated documents plus required attachments.

The wizard should not ask the planner for information the workbook already contains, and it should not force the planner to manually upload documents the wizard can generate.

---

# 2. Workbook Structure

Each wizard should use the calculation workbook as the source of truth.

The typical workbook structure is:

```text
Step 1 - Input Properties Info
Step 2 - Calculation Sheet
Case Info
```

or, depending on the workbook:

```text
Step 1 - Input Properties Info
Step 2 - Main Calculation Sheet
Case Info
```

## Step 1

Step 1 usually comes from GIS. Do not alter its labels, columns, layout, or naming conventions.

Common Step 1 fields include:

```text
OWNER_MAILTO
Mail_Addr_1
Mail_Addr_2
Mail_City
Mail_State
Mail_Zip
Total_Land_Area
group_dscr
TAX_ID
Legal_Dscr_1
Legal_Dscr_2
Legal_Dscr_3
Legal_Dscr_4
Property Description
Land Use
Lot Size
```

The templates can use clean placeholder names, but the source Step 1 tab must remain unchanged.

Example mapping:

```text
Template {{Owner_Name}} = Step 1 OWNER_MAILTO
Template {{Mailing_Address_1}} = Step 1 Mail_Addr_1
Template {{Mailing_Address_2}} = Step 1 Mail_Addr_2
Template {{Mailing_City}} = Step 1 Mail_City
Template {{Mailing_State}} = Step 1 Mail_State
Template {{Mailing_Zip}} = Step 1 Mail_Zip
Template {{Land_Area}} = Step 1 Total_Land_Area
Template {{Land_Use}} = Step 1 group_dscr
```

Blank second address lines should be suppressed automatically.

## Step 2

Step 2 holds the case calculation and staff analysis. Do not alter its official layout if that sheet becomes part of the RCA packet.

Step 2 usually provides:

```text
Application number
Date received
Date complete
Street name
Block/lot description
Cross streets
Side of street
Support percentage
SFR percentage
Minimum lot size or minimum building line
Qualifying lot/property count
Total lot/property count
Staff analysis table
```

Step 2 must be read carefully. The wizard should not export helper columns or formula rows into the packet.

## Case Info

`Case Info` is the wizard-control tab. It stores planner inputs and case-level fields that are not reliably in Step 1 or Step 2.

Examples:

```text
Letter_Date
Protest_Deadline
Planner_Name
Planner_Phone
Planner_Email
Council_District
Approval_Path
Council_Deadline
Applicant_Name
Applicant_Address_1
Applicant_Address_2
Applicant_City
Applicant_State
Applicant_Zip
Sign_Count_Number
Sign_Count_Words
Sign_Placement
Director_Name
Director_Approval_Date
Subdivision_Name
Subdivision_Plat_Year
Neighborhood_Age_Statement
Structure_Age_Statement
Preservation_Finding
PC_Signer
PC_Approval_Date
```

Use dropdowns wherever possible. Do not make planners type values that should be controlled options.

Important dropdowns:

```text
Approval_Path:
- Director Approval
- Planning Commission Approval

PC_Signer:
- Martha L. Stein, Chair
- Sonny Garza, Vice-Chair
```

For renewal workflows, additional fields may be needed:

```text
Original_Ordinance_No
Original_Effective_Date
Original_Expiration_Date
Application_Filed_Date
Application_Accepted_Date
Renewal_Complete_Date
Owner_Letter_Type
```

---

# 3. Mail Merge Logic

The wizard is essentially a mail merge engine with better control than Word mail merge.

For each document:

1. Load the official Word template.
2. Build a data dictionary from Step 1, Step 2, and Case Info.
3. Replace only the intended placeholders.
4. Preserve template formatting.
5. Save a new Word file.
6. Optionally generate a PDF proof.
7. Open the output folder after generation.

## Owner Letters

Owner letters use one row per owner/property from Step 1.

The wizard should support:

```text
Preview One Letter
Generate All Letters
```

Preview should use the first valid owner row automatically.

A valid owner row usually has:

```text
Owner name
Mailing address
Property information
```

The planner should not need to enter a row number.

## All-owner output

The combined owner output should be one Word document.

Example:

```text
MLS_All_Owner_Notification_Letters.docx
MBL_All_Owner_Notification_Letters.docx
```

For notification letters, the selected map should be inserted after each owner letter.

The structure should be:

```text
Owner 1 letter
Owner 1 map
page break
Owner 2 letter
Owner 2 map
page break
Owner 3 letter
Owner 3 map
no extra final blank page
```

Avoid trailing blank pages. Add page breaks only between owner packets, never after the final owner.

## Map handling

If the map is a PDF, convert it temporarily to an image, insert it into Word, then delete the temporary image. The output folder should not contain leftover PNG files.

The notification buttons should output only Word documents unless explicitly designed otherwise.

Correct behavior:

```text
Preview One Notification Letter -> one .docx
Generate All Notification Letters -> one .docx
```

Incorrect behavior:

```text
Preview One Notification Letter -> .docx + leftover PNG
Generate All Notification Letters -> .docx + leftover PNG
```

---

# 4. Template Replacement Rules

The most important engineering rule:

**Never replace paragraphs by rebuilding them from scratch if the template has formatting that matters.**

Instead:

1. Locate the placeholder text.
2. Replace text within the existing Word run when possible.
3. Preserve run formatting:

   * font family
   * font size
   * bold
   * italic
   * spacing
   * paragraph alignment
   * table structure
4. Clear placeholder red color only for inserted replacement values, unless the template intentionally requires color.
5. Do not globally set the whole document to Arial 12.

This matters because City templates often have letterheads, titles, tables, footers, and bold labels that must remain unchanged.

## What not to do

Do not apply one font to the entire document.

Example of a bad fix:

```text
Set all paragraphs to Arial 12
```

This can damage letterhead, title text, seals, footer text, and heading fonts.

## What to do instead

Apply formatting only to the inserted body text if required, and leave headers/letterhead untouched.

If a template has:

```text
CITY OF HOUSTON
Planning & Development Department
```

that letterhead should stay exactly as the template defines it.

---

# 5. Case Info UI Rules

The Case Info tab and wizard UI should be clear, direct, and planner-friendly.

Use visual color conventions:

```text
Yellow = planner must enter
Blue = auto-read/review values from workbook
Green = wizard rules/dropdowns
```

The wizard UI should show the case info back to the planner before generating documents.

The planner should be able to review:

```text
Application number
Application type
Street/block description
Side of street
Cross streets
Minimum lot size/building line
Support percentage
SFR percentage
Approval path
Planner info
Applicant info
Council deadline
```

Do not bury critical fields. The Case Info page should be a quality-control screen, not just a data table.

---

# 6. UI Design Rules

The wizard should feel like a modern desktop product, not a script.

Recommended steps:

```text
1. Select Files
2. Review Case Info
3. Review Owners
4. Generate Documents
```

The front screen should have clear sections:

```text
Required Setup
- Calculation workbook
- Boundary map
- Output folder

City Templates
- Owner notification letter
- Applicant sign letter
- Director approval
- Planning Commission approval
- Routing form
- Draft ordinance

Packet Attachments
- Subdivision plat
- Tax map
- HCAD records
```

City templates should be prefilled from the wizard’s internal `templates` folder, but each should still have a Browse button in case the City updates the template.

The output folder can default to Documents.

The wizard should open the output folder after a successful generation.

## Color identity

Use different sidebar colors to distinguish wizards:

```text
MLS Wizard = blue
MBL Wizard = green
MLS Renewal Wizard = another distinct color, such as purple or orange
```

---

# 7. Approval Path Logic

Most workflows need:

```text
Approval_Path:
- Director Approval
- Planning Commission Approval
```

If Director Approval:

```text
Use Director Approval template
Check Director Approval on routing form
Do not check Planning Commission Approval
```

If Planning Commission Approval:

```text
Use PC Approval template
Check Planning Commission Approval on routing form
Do not check Director Approval
```

For protested cases, use Planning Commission path.

The approval documents must be grouped together in the UI under:

```text
Approvals
- Generate Director Approval
- Generate PC Approval
```

---

# 8. Routing Form Logic

The routing form must preserve the official template.

For MLS new applications:

```text
Lot Size checked
SMLS Block Application checked
Original Ordinance unchecked
```

For MBL applications:

```text
Building Line checked
SMBL Block Application checked
Original Ordinance unchecked
```

For MLS renewals:

```text
Lot Size checked
SMLS Block Renewal checked
Original Ordinance checked
```

The topic text must align properly when it wraps. The second line of the location should align with the main topic text, not shift under the label.

Incorrect:

```text
The topic of Request: Special Minimum Lot Size...
(The 4800 block..., between Calhoun Road and
Coffee Street)
```

where `Coffee Street` shifts left.

Correct:

```text
The topic of Request: Special Minimum Lot Size...
                      (The 4800 block..., between Calhoun Road and
                      Coffee Street)
```

or an equivalent table-based alignment that visually keeps the wrapped text in line.

Remove red placeholder coloring from generated values, including punctuation such as parentheses and percent symbols.

---

# 9. Director Approval and PC Approval Logic

These are not freeform narratives. They are structured evaluation forms.

The wizard must use the uploaded template and replace only the case-specific text.

Do not recreate the approval form from scratch.

## MLS New Application Approval Fields

Typically:

```text
Application number
Application area description
SFR percentage
Support percentage
Minimum lot size
Qualifying lot count
Total lot count
Neighborhood age statement
Structure age statement
Preservation finding
Director name/date or PC signer/date
```

## MBL Approval Fields

Use building-line concepts:

```text
Minimum building line in feet
Qualifying property count
Developed property count
Building-line percentage
Support percentage
SFR percentage
Building-line preservation finding
```

## Renewal Approval Fields

Use renewal-specific criteria:

```text
SMLSB renewal includes all property within at least one blockface
More than 60% developed with SFR units
No protest or protest status
Prevailing lot size logic
75% of lots exclusive of corner lots
Lots not varying more than 10% from average
Renewal meets criteria
```

The renewal approval form is not the same as the new MLS approval form.

---

# 10. Draft Ordinance Logic

The ordinance generator should fill factual planning data only. It should not pretend to complete Legal’s work.

Fill:

```text
Application area
Minimum lot size or minimum building line
Support percentage
Qualifying count
Total count
Exhibit A property descriptions
Exhibit B boundary map
Exhibit C property list
```

Leave Legal-controlled fields alone:

```text
Ordinance number
Passed/adopted/approved dates
Effective date
Attorney name/date
L.D. file number
Legal edits to protest language
```

## MLS new ordinance

Use lot-size language:

```text
Special Minimum Lot Size Block
Section 42-197
lot size character
minimum lot size
square feet
```

## MBL ordinance

Do not reuse MLS language. Use building-line language:

```text
Special Minimum Building Line Block
Section 42-170
building line character
minimum building line
feet
front property line
```

## MLS renewal ordinance

Use renewal language:

```text
renewing the establishment
original ordinance number
original effective date
expiration date
application filed/accepted/completed dates
20-year renewal term
Special Minimum Lot Size Block Renewal
```

Renewal ordinances are meaningfully different from new application ordinances.

---

# 11. RCA Packet Builder

The RCA packet builder should ask only for documents the wizard cannot already access or generate.

For MLS/MBL new applications, the wizard can generate or access:

```text
Routing Form
Director Approval or PC Approval
Boundary map
Step 2 staff analysis export
Draft ordinance, if needed
```

The planner should upload:

```text
Subdivision plat
Tax map
HCAD records
```

HCAD records must support multiple PDFs.

Use:

```text
Browse Multiple
```

and allow Ctrl/Shift multi-select.

## Packet order

Typical packet order:

```text
1. Routing Form
2. Director Approval or PC Approval
3. Boundary Map
4. Staff Analysis / Step 2 export
5. Subdivision Plat
6. Tax Map
7. HCAD Records
8. Draft Ordinance, if included in packet
```

If the manager removes a form from the packet, remove it from both:

1. the UI, and
2. the packet builder logic.

Example: if Checklist is removed, remove the Checklist picker/button and remove it from the final packet.

---

# 12. Step 2 Staff Analysis Export

This is a major pitfall.

Do not export the entire Excel used range if it includes helper columns, formula rows, or thousands of blank cells.

The wizard should export only the meaningful staff-analysis range.

Rules:

1. Use the official print area if it exists and is correct.
2. Otherwise detect the last meaningful row and column.
3. Exclude helper columns to the right.
4. Exclude rows that only contain zeros, formulas, placeholders, or formatting.
5. Stop after the last real property/calculation row.
6. Export as PDF for the RCA packet.
7. Fit to page width.
8. Do not create extra blank PDF pages.

For MLS, the correct range may be something like:

```text
Columns A:H
Rows 1 through the final total/property row
```

For MBL, similar logic applies, but the sheet name may be:

```text
Step 2 - Main Calculation Sheet
```

Always visually inspect a proof.

Bad export symptoms:

```text
#NAME? cells
helper columns on the right
lots of zero rows
blank pages
tiny unreadable form
two or more pages of empty grid
```

---

# 13. File Output Rules

Each button should produce exactly what its label says.

Examples:

```text
Preview One Notification Letter -> Word file only
Generate All Notification Letters -> combined Word file only
Generate Applicant Sign Letter -> Word file
Generate Director Approval -> Word file
Generate PC Approval -> Word file
Generate Routing Form -> Word file
Generate Draft Ordinance -> Word file
Generate RCA Packet PDF -> PDF packet
```

Do not leave temporary files in the output folder.

Bad behavior:

```text
Button generates Word doc plus stray PNG
Button generates debug PDFs
Button generates temporary map images
Button creates multiple old version files
```

Temporary files should go to a temporary directory and be deleted automatically.

After a successful generation, show a success message and open the output folder.

If Word has a file open and Windows blocks saving, show a friendly message:

```text
The output file appears to be open in Word. Close the file or choose another output folder, then try again.
```

---

# 14. Dependency and Packaging Rules

The final user should not need to run Python manually.

The wizard package should include:

```text
run_wizard.bat
install_requirements.bat
requirements.txt
templates/
sample_input/
sample_output/
README.md
```

The wizard should use external templates in the `templates` folder, not hard-coded Base64 templates embedded inside code.

Why:

```text
External templates are easier to update.
The City can replace a template without editing code.
It avoids accidentally using old templates.
```

Avoid requiring complicated system-level installs like Poppler if possible.

For PDF map conversion, prefer a Python package that installs through requirements, such as PyMuPDF, rather than requiring the planner to install Poppler manually.

---

# 15. Common Pitfalls

## Pitfall 1: Recreating templates

Wrong:

```text
Build a new Director Approval form from scratch.
```

Right:

```text
Use the uploaded Director Approval template and replace only the variable fields.
```

## Pitfall 2: Global font changes

Wrong:

```text
Set the whole document to Arial 12.
```

Right:

```text
Preserve template font and only format inserted case values where appropriate.
```

## Pitfall 3: GIS labels changed

Wrong:

```text
Rename Step 1 GIS columns to nicer names.
```

Right:

```text
Leave Step 1 alone and map GIS labels behind the scenes.
```

## Pitfall 4: Step 2 export includes blanks

Wrong:

```text
Export full used range.
```

Right:

```text
Export only the meaningful staff-analysis form range.
```

## Pitfall 5: Asking for unnecessary uploads

Wrong:

```text
Ask user to upload routing form, approval form, staff analysis, and map when the wizard already creates or has them.
```

Right:

```text
Only ask for external attachments: plat, tax map, HCAD records, and selected boundary map.
```

## Pitfall 6: Missing buttons

Every module in the backend must have a visible UI button.

If a generator exists, the UI needs an obvious button for it.

## Pitfall 7: Hidden bottom buttons

The Generate page must be scrollable. On smaller screens, bottom buttons like Draft Ordinance or RCA Packet Builder must not be cut off.

## Pitfall 8: Wrong approval grouping

Director Approval and PC Approval belong together under:

```text
Approvals
```

not scattered across unrelated sections.

## Pitfall 9: Packet builder calls wrong function

Make sure each UI button calls the correct generator.

Example:

```text
Generate Draft Ordinance -> draft ordinance generator
Generate RCA Packet PDF -> packet builder
```

Do not accidentally call draft ordinance logic from the packet builder button.

---

# 16. Recommended Build Order for a New Wizard

For any new RCA wizard, follow this build order:

```text
v1: Case Info tab + field map + UI shell
v2: Owner notification preview/all-owner mail-out
v3: Applicant complete/sign letter
v4: Director Approval + PC Approval
v5: Routing Form
v6: RCA Packet Builder
v7: Draft Ordinance
v8+: polish, UI cleanup, edge cases, manager changes
```

This order works because each module builds on the previous data model.

---

# 17. Wizard Types and Key Differences

## MLS Wizard

Purpose:

```text
Special Minimum Lot Size Block new applications
```

Language:

```text
Minimum lot size
Square feet
Lot size character
Section 42-197
```

## MBL Wizard

Purpose:

```text
Special Minimum Building Line Block applications
```

Language:

```text
Minimum building line
Feet
Building line character
Front property line
Section 42-170
```

## MLS Renewal Wizard

Purpose:

```text
Renewal of existing Special Minimum Lot Size Block
```

Language:

```text
Renewal
Original ordinance
20-year term
Prevailing lot size
75% of lots exclusive of corner lots
No protest / protest status
Original ordinance required
```

Renewal is not just MLS with a checkbox. It needs separate templates, separate findings, and renewal-specific ordinance language.

---

# 18. How to Think About the Wizard

A good wizard should reduce planner workload without hiding important review points.

The planner should always be able to:

```text
Review case info
Review owner data
Preview one letter
Generate all letters
Generate forms
Build final packet
Open output folder
Replace templates if City templates change
```

The wizard should not make legal conclusions beyond the template. It should fill planning facts and leave Legal blanks alone.

The highest standard is:

**The generated document should look like the planner manually filled the official City template perfectly.**

---

# 19. MLS Renewal Wizard Addendum

The MLS Renewal Wizard follows the same core philosophy as the MLS and MBL wizards: use the official City templates exactly as provided and inject case-specific values from the calculation workbook and `Case Info` tab.

However, renewal cases have several unique workflow, document, and packet requirements. Renewal is not just a normal MLS application with a renewal checkbox. It requires separate templates, separate findings, separate ordinance language, and the original ordinance as an RCA packet attachment.

---

## 19.1 Renewal-Specific Workflow

The MLS Renewal Wizard is used for:

```text
Special Minimum Lot Size Block Renewal
```

The wizard should preserve the same major UI flow:

```text
1. Select Files
2. Review Case Info
3. Review Owners
4. Generate Documents
```

The sidebar color should be distinct from regular MLS and MBL. The preferred color for MLS Renewal is:

```text
Dark red / burgundy
```

The wizard should include renewal-specific document generation buttons:

```text
Owner Notification Mail-Out
- Preview One Renewal Notification Letter
- Generate All Renewal Notification Letters

Applicant Letter
- Generate Applicant Sign Letter

Approvals
- Generate Director Approval
- Generate Planning Commission Approval

RCA Packet Forms
- Generate Renewal Routing Form

RCA Packet
- Generate MLS Renewal RCA Packet PDF

Draft Ordinance
- Generate Draft Renewal Ordinance
```

The Draft Renewal Ordinance should be a standalone Word document. It should not be inserted into the RCA packet PDF unless the user or department later changes that requirement.

---

## 19.2 Renewal Owner Letter Types

The MLS Renewal Wizard has two owner notification letter types:

```text
SFR Renewal Letter
No-SFR Renewal Letter
```

The wizard should include an `Owner_Letter_Type` dropdown in `Case Info` or the UI:

```text
Owner_Letter_Type:
- SFR Renewal Letter
- No-SFR Renewal Letter
```

### SFR Renewal Letter

The SFR renewal letter includes land-use/map confirmation language. It asks the property owner to review the attached map and confirm the land use shown for their property.

The SFR letter should pull owner/property values from Step 1, including:

```text
Owner name
Mailing address
Total land area
Land use / group description
```

It should pull case-level values from Step 2 and `Case Info`, including:

```text
Letter date
Protest deadline
Planner name
Planner phone
Planner email
Block description
Cross streets
Minimum lot size
```

### No-SFR Renewal Letter

The No-SFR renewal letter is shorter and does not include the same land-use/map confirmation step. It focuses on:

```text
Renewal notice
Minimum lot size
Owner lot size
Protest deadline
Future notice if protest is filed
Planner contact information
Spanish notice
```

---

## 19.3 Spanish Notice Formatting Rule

Some renewal owner letters include a Spanish notice near the bottom of page 1.

The Spanish notice must remain on page 1 without damaging the rest of the template.

Correct approach:

```text
Reduce only the Spanish notice paragraph to 10.5 pt.
If needed, allow fallback to 10 pt.
```

Incorrect approach:

```text
Do not globally shrink the letter.
Do not compact the whole body.
Do not alter the City logo or letterhead.
Do not change the Next Steps indentation.
Do not change the date, subject, footer, or body font globally.
```

This is a narrow formatting exception only for the Spanish notice paragraph.

---

## 19.4 Renewal Approval Logic

Renewal approvals are different from new MLS approvals.

The wizard must use renewal-specific approval templates:

```text
MLS_REN_Director Approval
PC_Approval_NEW
```

The approval documents should be grouped in the UI under:

```text
Approvals
- Generate Director Approval
- Generate Planning Commission Approval
```

The wizard should keep `Approval_Path` as a planner-controlled dropdown:

```text
Approval_Path:
- Director Approval
- Planning Commission Approval
```

General business rule:

```text
No protest -> Director Approval
Protest filed -> Planning Commission Approval
```

However, the wizard should not silently override the planner’s selected approval path. If possible, it may show a warning if protest data conflicts with the selected approval path.

### Renewal Director Approval Findings

The Renewal Director Approval should use renewal-specific criteria, including:

```text
SMLSB renewal includes all property within at least one blockface
More than 60% of the proposed SMLSB renewal is developed with SFR units
No protest was received or protest status language
Renewal will further the goal of preserving the lot size character of the area
Prevailing lot size exists when 75% of lots, exclusive of corner lots, do not vary more than 10% from the average
```

### Renewal PC Approval Findings

The Renewal Planning Commission Approval should also use renewal-specific findings. It should not use the new MLS application approval language.

PC approval language may reference that evidence of support is not required by the ordinance and that property owners were notified of the protest period.

---

## 19.5 Renewal Routing Form Logic

The Renewal Routing Form must preserve the official routing form template.

For MLS Renewal cases, the routing form should check:

```text
LOT SIZE
SMLS Block Renewal
Original Ordinance
```

It should not check:

```text
BUILDING LINE
SMLS Area Application
SMLS Block Application
SMBL Block Application
SMBL Block Renewal
```

The topic/request block should wrap cleanly and align visually under the main topic text.

Correct visual behavior:

```text
The topic of Request:  Special Minimum Lot Size/ Block Renewal No. XXX
                       (xxxx block Street Name, north and south sides,
                       between Street A and Street B)
```

The final cross-street line should not drift left under the label.

---

## 19.6 Renewal RCA Packet Builder

The Renewal RCA Packet PDF should be compiled using generated documents and external attachments.

The Draft Renewal Ordinance should remain separate as a standalone Word document.

The Original Ordinance should be included as an uploaded PDF attachment in the RCA packet.

Renewal packet order:

```text
1. Renewal Routing Form
2. Renewal Director Approval or Renewal Planning Commission Approval
3. Boundary Map
4. Step 2 Staff Analysis PDF
5. Subdivision Plat
6. Tax Map
7. HCAD Records
8. Original Ordinance PDF
```

The wizard should ask the planner to upload only external attachments that the wizard cannot generate:

```text
Boundary map
Subdivision plat
Tax map
HCAD records
Original ordinance
```

HCAD records should allow multiple PDF selections.

The packet builder should not ask the planner to upload:

```text
Routing form
Director Approval
Planning Commission Approval
Staff Analysis
```

because the wizard can generate or export those.

---

## 19.7 Renewal Draft Ordinance Logic

The Draft Renewal Ordinance should be generated as a standalone Word document.

The wizard should use:

```text
MLS_RENEWAL_Ordinance Template.docx
```

as the source of truth.

The ordinance generator should follow the same mail-merge-style philosophy as the rest of the wizard:

```text
Use the uploaded ordinance template.
Replace only case-specific text.
Preserve paragraph formatting, font, spacing, sections, and legal structure.
Do not recreate the ordinance from scratch.
Do not rewrite Legal-controlled language beyond the factual planning fields.
```

The renewal ordinance may require additional `Case Info` fields:

```text
Original_Ordinance_No
Original_Effective_Date
Original_Expiration_Date
Application_Filed_Date
Application_Accepted_Date
Renewal_Complete_Date
Legal_File_No
Ordinance_Year
```

The ordinance generator may fill planning factual fields such as:

```text
Application area
Minimum lot size
Qualifying lot count
Total lot count
Percentage meeting prevailing lot size criteria
Exhibit A property descriptions
Exhibit B boundary map
Exhibit C property list
```

The ordinance generator should leave Legal-controlled fields alone unless the department specifically instructs otherwise:

```text
Final ordinance number
Passed/adopted/approved dates
Effective date
Attorney name/date
Final L.D. file number
Legal edits to protest language
```

---

## 19.8 Renewal Calculation Workbook Rules

The renewal calculation workbook remains the source of truth.

Do not alter:

```text
Step 1 - Input Properties Info
Step 2 - Calculation Sheet
Step 2 - Main Calculation Sheet
```

depending on the workbook’s naming.

Add or maintain a `Case Info` tab for renewal-specific planner inputs.

Renewal-specific `Case Info` fields may include:

```text
Owner_Letter_Type
Approval_Path
Letter_Date
Protest_Deadline
Planner_Name
Planner_Phone
Planner_Email
Council_District
Council_Deadline
Applicant_Name
Applicant_Address_1
Applicant_Address_2
Applicant_City
Applicant_State
Applicant_Zip
Sign_Count_Number
Sign_Count_Words
Sign_Placement
Director_Name
Director_Approval_Date
PC_Signer
PC_Approval_Date
Original_Ordinance_No
Original_Effective_Date
Original_Expiration_Date
Application_Filed_Date
Application_Accepted_Date
Renewal_Complete_Date
Legal_File_No
```

The wizard should pull renewal calculation results from Step 2 when possible:

```text
Application number
Block description
Cross streets
Side of street
Minimum lot size
SFR percentage
Total lots
Lots meeting prevailing lot size criteria
Percentage meeting prevailing lot size criteria
Staff analysis table
```

If Step 2 contains formulas, the wizard should use displayed/cached values where possible and avoid exporting formula-helper areas.

---

## 19.9 Renewal Step 2 Staff Analysis Export

The Renewal RCA Packet PDF should include a Step 2 Staff Analysis PDF export.

The export should capture only the meaningful visible staff-analysis area.

It should not capture:

```text
Helper columns
Formula-only rows
Blank grid pages
Placeholder-only ranges
Thousands of blank cells
```

Use the same rule as the MLS and MBL wizards:

```text
Use official print area if correct.
Otherwise detect the final meaningful row and column.
Fit to page width.
Export only the staff-analysis content.
Visually inspect the proof.
```

---

## 19.10 Renewal Output Rules

Each button should produce exactly what it says.

Correct renewal outputs:

```text
Preview One Renewal Notification Letter -> one Word document
Generate All Renewal Notification Letters -> one combined Word document
Generate Applicant Sign Letter -> one Word document
Generate Director Approval -> one Word document
Generate Planning Commission Approval -> one Word document
Generate Renewal Routing Form -> one Word document
Generate MLS Renewal RCA Packet PDF -> one PDF packet
Generate Draft Renewal Ordinance -> one Word document
```

The wizard should not leave temporary files in the output folder, including:

```text
Temporary PNG map images
Temporary packet PDFs
Debug files
Intermediate DOCX files
```

Temporary files should be created in a temporary working directory and deleted automatically.

After successful generation, the wizard should show a success prompt and open the output folder when the user clicks OK.

---

## 19.11 Renewal-Specific Quality Checks

Before a Renewal Wizard build is considered complete, verify:

```text
City logo/letterhead remains intact in owner letters and applicant letter
Spanish notice fits on page 1
Next Steps numbering resets for every owner
SFR and No-SFR subject blocks align correctly
Map appears after each owner letter
No extra blank page after the final owner packet
Director Approval uses renewal criteria
PC Approval uses renewal criteria
Routing Form checks SMLS Block Renewal
Routing Form checks Original Ordinance
Routing Form topic text wraps cleanly
RCA Packet includes Original Ordinance PDF
RCA Packet does not include Draft Renewal Ordinance unless explicitly requested
Step 2 Staff Analysis export has no blank helper pages
Draft Renewal Ordinance is standalone Word output
Draft Renewal Ordinance uses renewal language and preserves the uploaded template
```

Whenever a new Renewal Wizard version is delivered, include rendered screenshots or previews of changed documents so formatting can be verified before downloading the full package.

