# Changelog

## v13.7 - Step 2 / Case Info Source-of-Truth Audit
- Re-applied official Step 2 calculation values after all fallback calculations so Step 1 fallback logic can no longer override planner-visible Step 2 values.
- Fixed total-lot count for approvals, routing form, draft ordinance, and packet generation so EXC rows marked on Step 2 are excluded consistently.
- Added formula-cache fallbacks that still respect Step 2 EXC markings when Excel has not cached formula results.
- Routing Form now uses the same final Step 2/Case Info context as Director Approval, PC Approval, and Draft Ordinance.
- Verified MLS 905 exclusion scenario: SFR 70%, support 66%, total lots 20, qualifying lots 11.

## v13.6 - Step 2 SFR Source-of-Truth Fix
- Director Approval and PC Approval now use the visible Step 2 SFR percentage from `Step 2 - Calculation Sheet!H11`.
- Prevents fallback property-count logic from producing 67% when the official Step 2 staff analysis shows 64%.
- Keeps routing form, approval forms, and packet output aligned with the calculation sheet.

# v13 Change Log

- Removed RCA Checklist from the wizard UI.
- Removed RCA Checklist template picker from Select Files.
- Removed Generate RCA Checklist button from Generate Documents.
- Removed RCA Checklist from the RCA Packet PDF assembly order.
- Kept Routing Form, Approvals, Draft Ordinance, and Packet Builder functions intact.


## v13.1
- Fixed PDF map conversion for Draft Ordinance and other Word outputs. The wizard now uses PyMuPDF instead of requiring Poppler/pdftoppm.
- Updated requirements.txt to include PyMuPDF.

## v13.2
- Reworked support-percentage logic so Director Approval, Planning Commission Approval, and Draft Ordinance pull the visible support percentage from Step 2 of the calculation sheet.
- Step 2 displayed support percentage now overrides recomputed fallback support values.
- Kept Step 1 and Step 2 workbook layouts unchanged.

## v13.3
- Fixed Draft Ordinance font consistency issue where generated support/lot-size text could render smaller than the surrounding ordinance text.
- Draft Ordinance now normalizes missing run-level font metadata to match the City ordinance body font without changing bolding or spacing.


## v13.4
- Fixed MLS Routing Form topic/location paragraph wrapping so the cross-street line aligns with the location text instead of drifting left under the label.
- Kept template formatting and removed placeholder red styling from filled location text.


## v13.6 - Ordinance side-phrase template fix
- Updated the draft ordinance template so ordinance-specific grammar is handled in the ordinance output, not by forcing Case Info `Side_Of_Street` / B46 to include `of the`.
- Added ordinance-only component fields for the draft ordinance title and WHEREAS clause. Routing Form, Director Approval, PC Approval, and cover/subject fields continue to use the clean side phrase such as `north and south sides`.
- Added a rendered sample proof under `sample_output/render_ordinance_v13_5/`.
