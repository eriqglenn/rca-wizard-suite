# ILMS Holds

In all four wizards, click ILMS Holds beside Import Application PDF / Fill Calculation Sheet.
The current workbook supplies the application number and property addresses. Chrome opens
for the planner to sign in directly to ILMS. After sign-in, the wizard automatically finds
each exact situs address, adds the application hold and verifies it in the hold list.
No separate Load, Find or Place buttons are required. The wizard does not read passwords.

EXC properties are included. Split-lot rows at the same address share one hold.
Existing holds remain and do not prevent the new application hold. Missing or ambiguous
addresses are reported for correction; a different address is never substituted.
Stop pauses before the next submission. Start / Resume continues using the case log;
verified submissions are not duplicated and uncertain submissions are reconciled first.
The case output folder receives a CSV report and resumable JSON log.

If the browser disconnects or an address/navigation step fails, the batch pauses at
that property. Later addresses show Not attempted, rather than Needs review. Fix
the displayed issue, then click Start / Resume; Chrome reopens if it was closed.
Keep the same case output folder and its JSON log so completed holds are skipped.
Needs review is retained when an Add may have reached ILMS but was not verified;
the wizard checks that record before attempting any further submission.

Login: https://ilms.internal.example/Ilms_Client/login
The wizard types GE and 140 directly, selects Hard Hold, fills the reason and confirms the Add request.
Live workflow verified on September 24, 2026: one hold for application 121 at 12107 Crystalwood Drive.
The live check used the actual ILMS screens; automated browser coverage uses local fixtures.
