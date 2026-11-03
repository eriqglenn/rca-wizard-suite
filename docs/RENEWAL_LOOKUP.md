# Previous designation lookup

MLS Renewal and MBL Renewal look for the previous designation when you save an imported application’s selected properties or create a boundary or tax map. No AI or extra approval screen is required.

The wizard matches the selected parcel shapes against the City’s published designation boundaries. It fills blank Original Ordinance No, Original Council Passage Date, Original Effective Date and Original Expiration Date fields when the corresponding value is verified. Passage is kept separate from effective date. A map start date is never substituted for either date.

Case Info records the lookup sources, notes and check time. The importer also retains the complete lookup in its existing source record. Existing planner values and formulas are preserved. Conflicts, missing dates and multiple matches appear as review notes; maps and calculation sheets remain available. For multiple prior designations, enter the intended original ordinance number in Case Info and refresh the map.

Sources are the City’s MLS/MBL map records, the status workbook linked from its public program page, and published Council minutes near a known map start date when needed. Some older MBL records have no effective or expiration date. Those remain blank for entry from the original ordinance; the wizard does not calculate an assumed term.

Public records are cached for one day. Refresh parcel data also refreshes the designation lookup. If a source is unavailable, a saved source can be used with a dated warning. Location matching happens locally against citywide data. Original workbooks are backed up before an update. If Excel holds a workbook open, an updated copy is saved and selected.
