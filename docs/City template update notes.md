# City template update September 29 2026

The templates downloaded from Neighborhood Planning SharePoint are the source for this update to V26work, V30 and V35. Updated copies preserve the City wording and layout while replacing case placeholders and historical examples with workbook fields. Applicant and owner letters whose wording already matches remain in place.

The MLS renewal PC approval has been adapted to lot size and square feet. The MBL Director routing form has been adapted to building line and feet, as requested. Director and Planning Commission routing/OneMeeting variants are selected for the applicable case path. The MBL renewal Director form retains the new source's 60% wording. The MBL renewal ordinance reserves Exhibit A for Legal and uses the selected case map instead of the historical source map.

## Where the values come from

Application number, boundary and proposed minimum come from the existing calculation readers. Lot counts retain separate rows for split lots and exclude EXC rows. Percentages describing lots use lot counts. Percentages describing the application area use the sum of lot areas. HCAD accounts are not deduplicated when counting split legal lots.

Case Info supplies planner/contact, district/councilmember, applicant property address, subdivision, protest count, hearing date and deadlines. Existing values in the calculation sheet take precedence where the wizard already reads a calculated result.

Additional source fields use the following Case Info names. In older workbooks, add the field name in column A and its value in column B. Missing substantive information prints a visible completion marker; the wizard does not invent it.

| Field | Meaning |
|---|---|
| Original_Passage_Date | Actual adoption/passage date of the original ordinance. The original effective date is separate. |
| PC_Noncorner_Lot_Count | Planner-confirmed count, excluding corner lots, described by the renewal PC approval. |
| Subdivision_Plat_Year | Year the subdivision was platted. |
| Structure_Age_Statement | Sentence describing the age of the houses/structures. |
| PC_Hearing_Date | Hearing date used in the OneMeeting PC narrative. |
| PC_Approval_Date | Date next to the PC signature, if known. |
| Director_Approval_Date | Date next to the Director signature, if known. |
| Council_Meeting_Date | Council agenda meeting date, if assigned. |
| PC_Agenda_Item or PC_Agenda_Number | Agenda number. |
| PC_Presenter | Presenter name. |
| PC_Motion_By | Commissioner making the motion. |
| PC_Second_By | Commissioner seconding the motion. |
| PC_Effective_Date | Planning Commission effective date. |

Unassigned meeting/signature fields remain available for completion. Signature and Legal fields are not filled with guessed dates or signatures.

## Source exception

The two downloaded MBL ordinance copies are protected and could not be read. Their installed ordinance template remains unchanged. The downloaded originals and the pre-update templates are retained in the maintenance audit folder.
