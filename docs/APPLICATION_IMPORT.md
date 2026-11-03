# Import an application

In MLS, MBL, MLS Renewal or MBL Renewal, select **Import Application PDF / Fill Calculation Sheet** under Required setup. The importer creates a calculation sheet for the wizard you opened.

1. Choose the application PDF. The scan appears next to editable suggestions. Review names, dates, the street, block or address range, sides and both cross streets. Handwriting and crossed-out entries often need manual entry. The application signed date and received date are separate.
2. Check the application review box and choose **Find HCAD Parcels**. Review the blue parcels and the nearby gray parcels. Corner properties, vacant tracts and multiple accounts at one address may need attention. Click a parcel on the map to locate its row. Include or remove rows as needed; **Add HCAD account** finds an additional parcel.
3. Use **Edit property** to review reported lot area, land-use classification, owner mailing information and legal description. A correction requires a short source note. `EXC` remains in Step 1 and in the map boundary. New applications have a reviewed Y/N support field; renewals have their original/current property-value fields instead.
4. **Download HCAD PDFs** searches the selected account numbers through the existing Chrome downloader. Complete HCAD's human verification when requested. **Read HCAD PDFs** accepts records already downloaded. PDFs are matched by the account printed inside the record. Recognized land-area differences appear in the review notes; they do not silently replace a reviewed value. Open an attached PDF from Edit property to confirm it.
5. Accept the boundary and selected properties using the parcel review box, then select **Save New Calculation Sheet**. Incomplete application fields, unconfirmed lot areas/classifications and edited boundary wording produce review notes instead of blocking the save. The accepted parcel selection is retained; uncertain numeric values and dates stay blank for completion in Excel. Choose a new filename in the case folder. The importer preserves the original application and calculation template. It creates the workbook and a matching `_sources` folder containing the application, attached HCAD PDFs, original submitted values and **Review notes.txt** when there are warnings. Saving does not certify that the calculation is complete; review Step 2 after completing missing facts. A writable destination, at least one selected parcel, valid distinct HCAD accounts and sufficient template capacity are still required to create a usable file.
6. The new workbook is selected back in the same wizard. Click **Load Workbook** to recalculate and review Step 2 using desktop Excel or LibreOffice. Then use **Create Boundary Map** and the normal letter tools.

## What differs by wizard

| Wizard | Additional information |
| --- | --- |
| MLS | Current lot areas go into Step 1. The City's Step 2 formulas calculate the minimum lot size. |
| MBL | Enter each property's building-line measurement in **Edit property**, or complete Step 2 column E after saving. The City's formulas determine the minimum. |
| MLS Renewal | Review the **Minimum to renew (square feet)**. Original lot areas can be entered in **Edit property**; current areas come from HCAD. |
| MBL Renewal | Review the **Minimum to renew (feet)** and each property's original/current building line in **Edit property** or Step 2. |

Renewal minimums are written to Step 2 D2. A clear, consistent labelled minimum or attached boundary-map caption can supply an editable suggestion. Current HCAD lot area is never used to infer an existing designation or setback. Missing historical values and building-line measurements stay blank, with a note explaining what to complete. The template's example measurements are removed.

Renewal application numbers accept an optional REN suffix and are saved with one REN suffix. Original ordinance number, effective/expiration dates and application filed/accepted dates have separate review fields. Enter details that are not reliably available from the scan. The map continues to use the selected wizard's lot-size/building-line and renewal wording.

## Data and review

Application scans are read locally using Windows text recognition. Both earlier forms and the City's 2026 layout are supported. Printed example boundaries are ignored; main-applicant fields are kept separate from co-applicant fields. If a PDF appears to be a different application type, a notice identifies the difference while allowing the planner to continue testing. No AI account or paid API is required. On a computer without an available OCR language, the PDF preview and manual fields remain usable. If local policy prevents the Windows OCR helper from running, use the preview and enter the application details manually.

HCAD PDFs with unreadable embedded text are also read locally. Small printed account numbers receive a closer text-recognition pass. Search-result and verification pages cannot be attached as property records merely because their filenames contain an account number. Check the attached record, especially where a lot-area difference is reported.

The county HCAD parcel layer supplies account, owner, mailing and site information. The City's HCAD layer supplies its explicit Total_Land_Area field and additional property details where available. These layers can have different tax years. The source year and any differences are retained per property; neither layer is assumed to be the latest individual HCAD record. Never substitute polygon area for the appraisal record's land area. Verify uncertain values against the HCAD PDF.

HCAD accounts retain their leading zeros in searches, Step 1 and map matching. If a public record has an invalid or missing account, the remaining records still load. The warning identifies the record to review; confirm its account and use **Add HCAD account** if it belongs in the boundary. Unresolved source records remain in the saved import review. Large review areas use smaller City lookup requests so all parcels can receive reported lot areas.

HCAD property classifications are suggestions for the City's SFR, MF, COM, VAC and EXC categories. Review the actual use and any calculation exclusions. Signature pages and a larger property roster do not automatically expand the application boundary or establish support.

The current MLS template supports 109 property rows; MBL and both renewal templates support 83 rows in their calculation ranges. New-application area sums cover every supported row in imported copies. The City's calculation methods and formulas are preserved. An exact address range such as 4806-4834 remains in the source record; the map description uses its 4800 block. Multiple blocks use a hyphen.

Applicant phone/email and the signed application date are retained in the source record because the current City Case Info tabs have no matching contact fields. Missing values are not copied into unrelated planner fields. Case Info fields are matched by name because their row positions differ between templates. New imports clear earlier map, tax-map, plat, original-ordinance and HCAD selections so attachments from another case are not reused.
