# Peer Reviewer

Select **5. Peer Reviewer** in the left menu, directly below Generate Letters / Generate Documents. It is available in all four wizards without first generating or loading documents.

Choose the final RCA packet, HCAD records (individual PDFs or a combined PDF), draft ordinance, and calculation sheet. The calculation/HCAD selections and known output filenames are suggested from the open case. They can be replaced, including with a case created outside the wizard. PDF and DOCX documents are read directly. Legacy DOC files require the existing Word/LibreOffice converter. DOCX references use paragraph numbers, while PDF references use page numbers.

Select **Run Peer Review**. Results are grouped as Errors, Needs review, and Passed. Select a finding for its evidence and locations. **Open selected source** opens its file. After saving corrections in the source files, select **Run Again**. **Save Review Report** exports a self-contained HTML report that can be opened and printed in a browser, including to PDF.

The reviewer snapshots saved files, records their SHA-256 fingerprints, and does not change or recalculate them. Unsaved edits in Word/Excel are not reviewed. Source changes during review are flagged. Changing the selected files invalidates the displayed results. Cancel or leaving the page cancels background reading.

## Checks

- Workbook errors, missing cached results, property counts, eligible areas, and Step 1/Step 2 area differences.
- SFR share by lot count and support share by area remain separate. New-application minimums are independently recomputed under the worksheet's recognized cumulative threshold method; this is arithmetic verification, not approval of the applicable legal rule.
- Renewal unchanged-or-meets flags remain distinct from lots that meet the minimum. MBL feet are kept separate from property square feet.
- HCAD accounts, combined land areas for split rows, listed legal lots, addresses, owners and mailing addresses. Different HCAD versions and ambiguous descriptions require review. Repeated accounts for distinct lot rows are valid.
- Finished-document minima, qualifying counts/denominators, words versus numerals, application references, explicit SFR/support statements, routing counts, council district, prior ordinance references and effective/expiration dates.
- Unresolved merge placeholders and unreadable/missing evidence. Draft adoption/signature blanks are not automatically errors.
- Map legends and address/street text inside the actual RCA, source name, and readable HCAD-account coverage in the packet.

## Items that remain human review

OCR can misread small, rotated or handwritten text. Local Windows OCR is available as an option and its comparisons are marked **Needs review**, never confirmed passes/errors. The checker does not use Ollama, another language model, or an external upload service.

A flattened map does not establish parcel geometry. Confirm the contiguous boundary, excluded properties, label placement, actual legal lot divisions, and the yellow tax-map outline visually against the relevant records. Missing labels are not assumed absent solely because OCR missed them.

Signed support, legal eligibility, building-line measurements, corner-lot exclusions, recorded plat facts, and the correct vintage of HCAD records need their supporting evidence. Equal area allocation is an accounting check, not a survey measurement. The report never certifies a case as approved or prevents document generation. A Passed finding applies only to the specific comparison described.
