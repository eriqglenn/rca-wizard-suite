# Find Subdivision Plats

All four wizards have **Find Subdivision Plats** in Required setup.
It suggests subdivision names from the saved calculation sheet, then searches
the Harris County Clerk's Maps and Condos index. No completed minimum-lot-size
or building-line calculation is required to search. You can also type a name,
file number, or volume and page directly.

HCAD section/replat abbreviations may differ from the County index. The finder
suggests a broader name, while preserving the worksheet's exact name in the
subdivision list. Results are candidates: compare the name, section, date,
file number and volume/page before choosing a recording. Repeat the search for
each subdivision or replat that applies. Narrow a search if only part of a
large result set is returned.

**Download Selected Plat** opens a dedicated Chrome or Edge window. The County
requires a free Clerk account to view recorded images. Sign in there, then
open/download the complete PDF in the County viewer. The wizard captures the
PDF and shows its pages for review. It does not enter passwords, solve human
verification, buy copies, remove watermarks, or substitute a newly drawn map
for the recorded plat. The normal browser profile keeps the County session
locally at LocalAppData/RCAWizardSuite/CountyClerkBrowser, outside the release
package. Closing the finder closes only its dedicated browser.

If the County viewer cannot supply a PDF directly, use **Open County Website**,
save the document there, and choose **Add Downloaded Plat**. Downloaded PDFs can
also be added without signing in through the wizard. Check any page-count
warning and obtain missing pages before using the plat.

Use Previous page / Next page to review all pages. Select the applicable PDFs
in Downloaded plats; Ctrl-click changes individual selections. **Save Selected
Plats As** exports them where you choose. **Use Selected Plats** saves a PDF
in the case folder and selects the Subdivision plat attachment for the RCA.
Multiple selected PDFs are combined without cropping or removing pages. A
single PDF is preserved byte-for-byte. Unique names preserve earlier files.
The calculation sheet is read only; maps and other attachments are unchanged.

Internet is needed to search and retrieve County records. County-account
recovery is handled on the County website; never enter the password in the
wizard or a support message. The signed-in County download path must be
checked with a working account when available; the search and PDF import
paths can be checked independently.

Developer: RCA Wizard Tools.exe platcheck --out <folder> --ui --browser --live
checks search parsing, workbook reads, PDF preservation, each wizard's controls,
and download capture against a local PDF server. --live checks the public
County search; it does not sign into a County account.

Official sources:
- https://www.cclerk.hctx.net/Applications/WebSearch/MC.aspx
- https://www.cclerk.hctx.net/Applications/WebSearch/Registration/Login.aspx
