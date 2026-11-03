# MBL Field Map v2

## Step 1 - Input Properties Info
- OWNER_MAILTO -> owner name/address block
- Mail_Addr_1, Mail_Addr_2, Mail_City, Mail_State, Mail_Zip -> mailing address

## Step 2 - Main Calculation Sheet
- D4 -> block description
- B5/D5 -> cross streets
- B6/D6 -> side(s) of street
- L16 -> minimum building line feet

## Case Info
- Letter_Date
- Protest_Deadline
- Planner_Name
- Planner_Phone
- Planner_Email

## MBL-specific language
- Special Minimum Building Line Block Application
- Section 42-170
- Minimum building line in feet

## v8 additions
- Step 2 -> Case Info overlay (engine reads live Step 2 values; Case Info auto-cells are Step 2 formulas).
- HCAD: Step 1 Site_addr_1 + TAX_ID exported to HCAD_Records/MBL_<app#>_HCAD_Address_List.txt.
