# Office PDF conversion and work-computer diagnostics

September 7, 2026 — suite v1.17

The reported MBL renewal RCA failure occurred at conversion of the generated
routing form to PDF. In v1.16, `office_to_pdf` first called `word_available`,
which required both pywin32 and a particular App Paths registry entry. If that
check failed, it never attempted Office and discarded the distinction between
missing runtime support and missing Office. The log alone does not establish
which condition occurred on the work computer.

Changes:

- Treat App Paths and COM registration in both registry views as display hints.
  Actual conversion always attempts the selected Office application. Recalculation
  similarly attempts Excel without requiring a registry hint.
- Use Word for document export and Excel for workbook export independently.
- Report the actual application, exception type/message, engine preference and
  guidance to the PDF setup check. Do not automatically install dependencies or
  alter Office registration or workplace security settings.
- Add `core/pdf_setup.py` and the `pdfcheck` companion command. The home-screen
  **Check PDF setup** button runs two isolated workers with a 60-second timeout
  each. Each writes a harmless temporary sample, exports it through Office, and
  verifies that the PDF has pages. Reports remain local until the user copies or
  saves them. No case files are read by this check.
- Both code lines use the same converter and tests. The spare now has the same
  Word/Excel fallback and Windows pywin32 requirement as the packaged line.

Validation commands:

- `python -B tools/test_office_conversion.py`: 11 checks including missing registry
  hints, Excel without Word, error preservation, forced engine selection,
  32-bit COM registration, recalculation, worker timeout and the home-screen report.
- `python -B tools/test_workflow.py`: six existing planner-workflow checks.
- `python -B tools/validate.py`: 637 existing document-structure checks. These
  do not by themselves test desktop Office PDF conversion.

The desktop Office integration check separately generates a complete MBL renewal
RCA using the bundled sample, including Step 2 analysis, through Microsoft Office
only while forcing the old registration hints to be absent. Final execution
results are recorded beside the distribution ZIP.

References: [Microsoft Word automation](https://learn.microsoft.com/en-us/office/vba/api/word.application)
and [Windows registry views](https://learn.microsoft.com/en-us/windows/win32/sysinfo/32-bit-and-64-bit-application-data-in-the-registry).

The user's work computer must still run the setup check and verify its own RCA.
A successful development-computer export does not establish its Office installation,
automation permissions, licensing or case-specific conversion behavior.

Source validation completed September 7, 2026: 637 document checks, six
workflow checks and 11 Office checks passed on both code lines. Desktop Word
and Excel each created a readable one-page test PDF. The full MBL renewal
Office-only integration produced an eight-component, eight-page RCA with the
old registry hints forced absent. The sandbox logon session could not start
Office; these real Office tests succeeded in the normal desktop user session.
