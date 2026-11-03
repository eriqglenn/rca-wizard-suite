# MLS Planning Commission draft reports

In the **MLS wizard**, open **4. Generate Letters** and find **Planning Commission Draft Report**. These options are available in V30 only.

1. Choose **Generate PC Draft Report** to review the report details. Step 2 supplies the boundary, proposed minimum, lot counts, support and land-use figures. Case Info supplies the meeting date, protests, plat year and history when entered.
2. Confirm the agenda number, meeting date, written-protest count and narrative. Check the public-notice box only after notice has been transmitted.
3. In **Properties and protests**, double-click the applicable properties to mark protests on the map. Support comes from Step 2. A protest count alone does not identify which properties to mark.
4. Choose **Generate PC Map**. Review the land-use colors, blue support arrows, red protest crosses, gray hatching below the proposed minimum, and black boundary. The complete City PC legend always appears, including turquoise hatching for transportation and utility properties. Drag labels to move them; right-click street names to resize or rotate them. Use **Download Map** for PDF/PNG or **Use This Map** to save it and select it for the report.
5. In **Attachments**, select the application PDF and regular boundary map. The saved PC map is selected automatically. The calculation analysis can also be included.
6. Choose **Generate PC Draft Report** at the bottom. The output folder receives an editable Word report, a PDF when desktop Office is available, and review notes. Attachments appear in calculation / PC map / application / boundary-map order. Narrative stays editable; PDF attachment pages are inserted as images. The PC map is included upright on a landscape page.

The report uses separate calculations for:

- SFR percentage: SFR count divided by eligible lot count.
- Petition support: land area of eligible supporting properties divided by eligible total land area.
- Lots meeting the minimum: all eligible lots at or above the proposed size, including every tie, and their combined share of eligible area. This is not the ranked row where the cumulative table first reaches 70%.
- Boundary count: all listed properties, including excluded properties. EXC properties remain on the map but are excluded from the above eligibility calculations.

The reader does not change the calculation workbook. Missing narrative or unknown figures remain visible review placeholders. Map and report warnings do not prevent saving. Reopen the PC work area after editing the workbook to read the latest saved values; settings are associated with that workbook snapshot.

Blue support arrows point toward the application street. They follow the street geometry when the map is rotated and point inward from either side of the street. The legend arrow points downward. If the street geometry is unavailable, the preview warns that arrow direction needs review.

The PC map is separate from the regular boundary and tax maps. The detailed **group_dscr** in Step 1 supplies application-property colors; Step 2's broad SFR/MF/COM/VAC/EXC buckets continue to control calculation eligibility. Office, for example, can be COM for calculations and pink on the map. EXC does not mean a property is undeveloped. The public City parcel layer supplies detailed surrounding land uses and fills missing worksheet classifications where the account and parcel footprint match. Its tax year is shown on the map. Missing detailed data falls back to an available worksheet category; unclassified parcels stay white with a warning. Land-use colors and patterns match the planner's full PC legend, including all ten land uses even when absent from a particular view.

Duplicate accounts or descriptions without a separate matched parcel are listed in review notes, without inventing parcel divisions. A selected older map is still attached, with a warning when its saved workbook or protest markings differ. Regenerate previously saved PC maps to apply the updated legend and palette.
