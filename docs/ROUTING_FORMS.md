# Routing forms stay on one page

All four wizards keep the RCA routing form on a single Letter page. The two
checklists use fixed columns instead of Word newspaper sections, so Tax Records
and Original Ordinance stay with the other attachments. The original form text,
fonts, signature lines, and editable checkbox controls are retained.

Standalone routing documents and the routing form generated for an RCA packet
use the same layout. If an Office converter still produces extra routing pages,
the packet builder fits all their content onto one page before appending the
approval and attachments. It never drops the overflow page or clips off its text.
Only the routing form is affected; approval documents and attachments keep their
own pagination.

Developer check: RCA Wizard Tools.exe routingcheck --out <audit folder>
Add --engine libreoffice or --engine msoffice to check rendered page counts;
--engine both checks both local Office converters. --workbook can point to the
MBL 91 renewal case for its regression check.
