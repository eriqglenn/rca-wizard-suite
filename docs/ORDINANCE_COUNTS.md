# Draft ordinance lot counts

MLS now fills the count in words and digits automatically when the Case Info
Lots_Meeting_Minimum field is blank or absent. It counts each eligible Step 2
property whose lot area is at least the proposed minimum shown on Step 2.
Every tie at the minimum is counted. Separate property rows are counted even
when they share an address or owner. EXC properties remain on the boundary map
but are excluded from this calculation, consistent with the Step 2 total.

A nonnegative whole number in Lots_Meeting_Minimum remains a planner override.
Leave it blank to calculate automatically. The same count is used by MLS
approvals and routing forms that share the calculation context. Neither the
workbook nor its formulas are changed by document generation.

Save/recalculate the workbook in Excel before generating documents. If the
minimum or eligible property data cannot be read, the count remains X (X);
document generation is not blocked. An invalid manual count also stays X (X)
until corrected or cleared. Zero is printed as zero (0).

MBL already counts developed properties meeting its minimum building line;
its number words now support counts above 30. MLS Renewal and MBL Renewal
ordinances have no qualifying-count clause and retain their renewal wording.
The old MBL template marked REVIEW_NEEDED is not the default ordinance.

Developer check: RCA Wizard Tools.exe ordinancecheck --out <audit folder>.
