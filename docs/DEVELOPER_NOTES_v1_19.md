# Follow-up audit: file selections and hidden errors

September 7, 2026 — suite v1.19

The user reported that fixes were not reliably working and requested a deeper
check. No new v1.18 work-computer traceback was supplied during this audit.
The last available work-computer traceback remains the v1.17 failure deleting
an old `_renewal_packet_work` directory. It is not sufficient to infer an Office
installation failure, a v1.18 failure, or which launcher was subsequently used.

## Reproduced before changing code

A fresh extraction of the delivered v1.18 ZIP was launched in the normal desktop
session, with its own temporary user profile/session and copied sample inputs.
Only Windows/System32 were on PATH, and Microsoft Office was the forced engine.
Using the actual packaged window, a different output folder was pasted into
Select Files, followed by Generate Documents and Generate MBL Renewal RCA Packet
PDF. The packet succeeded, but was written to the PREVIOUS folder. The pasted
folder had no PDF. This reproduced a user-visible failure which the previous
637 document checks and source packet-integration test did not cover.

The three dictionary-based wizards created editable Entry widgets without a
textvariable or edit handler. Only the Browse callback changed `self.files`.
Consequently typing, pasting or clearing any path left the previous internal
value active. This affected workbook refresh, attachments, template overrides
and output folders. MLS already used bound StringVars and did not have this bug.

A second regression reproduced an MLS packet error callback throwing NameError:
Python clears the exception variable after leaving `except`, but the queued
lambda still referenced that variable. It could hide the original conversion
failure. The same pattern occurred in eight MLS document actions.

## Changes

- Bind each affected file Entry to a retained StringVar and update `self.files`
  on every write. Browse and manual edits now use the same state. Page changes
  retain the current selection, and clearing a field clears its active value.
- Capture each MLS exception as a lambda default before queuing the callback.
- Include suite version, source/EXE launch mode and runtime path in the shared
  error log, while retaining the original traceback and non-throwing logger.
- Apply the same fixes to the spare source release v1.10.8.

There is no new Office converter, template, formula or ordinance change in this
release. The prior Office detection and independent temporary-workspace fixes
are retained. These proven defects do not by themselves establish the cause of
any unreported failure on the work computer.

## Verification

`tools/test_file_selection.py` has seven tests. The first six, before their fixes, produced
12 failing subtests and one error; all six passed after the fixes on both lines.
The tests exercise real Tk widgets, navigation and workbook switching, typed
attachments and output during RCA assembly, Browse/typing/session interactions,
delayed MLS failures, and build/runtime metadata. Office PDF export is controlled
with PDF fixtures in this regression suite; live Office is tested separately.

Existing checks: six workflow tests, 11 Office tests, five packet-workspace tests,
and the 637 document checks. Final results and the separate test through the
rebuilt packaged GUI are recorded in the validation file beside the main ZIP.
The packaged UI test uses copies of bundled samples and isolated settings; it
does not change the planner's cases or normal session settings.

A development-computer pass is not proof of the work computer's Office setup.
A current error traceback or Check PDF setup report is still needed to identify
any remaining work-computer failure. Do not describe this release as resolving
an unknown work-computer error.

A seventh test reproduced a cleared output folder becoming the current working
directory. The affected document actions now validate their output folder
before generation. All three affected wizards pass that regression.
