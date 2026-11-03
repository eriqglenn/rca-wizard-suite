# RCA packet retries with locked working folders

September 7, 2026 — suite v1.18

The newest work-computer log (v1.17, September 7 at 14:40 and 14:44) stops at
`shutil.rmtree(work)` in the MBL renewal packet builder. Windows returns
`PermissionError [WinError 5]` for the existing `_renewal_packet_work` directory.
This happens before any Office PDF conversion. The log does not identify which
process or permission prevents deletion, and does not establish whether Office
conversion now works on that computer. The older v1.16 entries describe the
separate Office detection problem addressed in v1.17/v1.10.6.

## Implementation

`core/packet_work.py` provides a context manager which creates a uniquely named
`rca_packet_...` child of the resolved system temporary directory for each attempt.
All four active packet builders and both older MLS/MBL engine entry points use it.
No attempt deletes or reuses a previous case-folder workspace. Each final RCA
still saves to the planner's chosen output path.

The context manager cleans up its own new workspace on success or failure.
Before recursive deletion it verifies that the resolved target remains the
original direct child of the system temp directory and is not a symbolic link.
An `OSError` during cleanup is ignored; a leftover temporary folder is preferable
to hiding the original conversion failure or reporting a completed packet as
failed. No later packet attempt depends on deleting that folder. No case-folder
permissions, Office registration or security settings are changed.

The new module is explicitly included in both packaged executables on the main
line. The same source change is applied to the spare v1.10.7 line. City templates,
instructional examples, workbook formulas, document wording and packet order are
unchanged.

## Validation

Both source lines pass:

- `python -B tools/test_packet_work.py`: five tests covering unique simultaneous
  workspaces, cleanup access denial, preservation of the original error, two
  generations in each of the four wizards while old case folders are blocked,
  and successful packet completion despite cleanup denial. Packet assembly uses
  real generated Word documents with controlled PDF conversion fixtures.
- `python -B tools/test_workflow.py`: six existing workflow tests.
- `python -B tools/test_office_conversion.py`: 11 existing Office tests.
- `python -B tools/validate.py`: all 637 document-structure checks.

A separate Windows integration test held a file handle without delete sharing
inside an old `_renewal_packet_work` fixture. An attempted deletion confirmed a
real Windows sharing violation (WinError 32). With that handle still open, two
complete MBL renewal sample RCAs were generated through desktop Word and Excel,
including Step 2 analysis. Both produced eight components and eight PDF pages;
the old folder and sentinel were unchanged. Unit tests separately simulate the
reported access-denied error (WinError 5). Actual Office exports ran in the normal
desktop session because the sandbox logon session cannot start Office here.

These results establish behavior on the development computer. The work computer
still needs to retry its RCA in this version. If conversion then fails, use
**Check PDF setup** and preserve the new error log. Final package integrity and
packaged validation results are recorded beside the main distribution ZIP.
