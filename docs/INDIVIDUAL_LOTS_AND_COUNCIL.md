# Individual lots, platted divisions and council information

Available in V30, V35 and V40. These features use records and programmed rules, without AI.

## Individual lots in Step 1

When importing an application, select the application properties and choose **Review individual lots**. Descriptions such as `LTS 1 & 2 BLK 4` or `LTS 1-3 BLK 4` are recognized. Each lot receives its own Step 1 row with the shared HCAD account and ownership information.

Enter the area and land use for each lot and identify the source. Use **Equal areas** only after verifying equal sizes against a plat or other reliable record. The individual areas must reconcile with the property's total. Partial lots, tracts and exceptions need manual review. Unknown areas may be saved explicitly as a draft; complete them before relying on the calculation results. Historical lot sizes and building-line measurements need their own verification.

For an existing Excel sheet, choose **Review Individual Lots** on the wizard's setup page. Save a new calculation copy, then use **Load Workbook** to recalculate Step 2. Already separated lots remain separate; the wizard does not split them again. The City columns, formula layout and artwork are preserved. A hidden `_Wizard Lots` sheet stores the lot identities and review information. Keep that sheet when transferring the workbook.

Notifications continue to be produced per underlying property. Separating several lots under the same HCAD account does not create duplicate property notices. Different properties with the same owner or mailing address still receive separate notices.

## Actual lot divisions on maps

**Show tax-parcel boundaries** controls the existing property outlines. **Platted lots...** opens the new editor for the boundaries inside those properties. This editor is available on boundary and PC maps; tax maps have **Platted lot divisions...**.

1. Complete and review the individual lot rows in Step 1, then create the map.
2. Open the platted-lot editor and select a property.
3. Load the relevant plat PDF or image. Select the correct page. Use the mouse wheel to zoom and right-button dragging to move around the plat.
4. Choose **Align plat: 3 points**. Click a recognizable corner on the parcel map, then the same corner on the plat. Repeat for three points that do not lie on one straight line. Check other corners to confirm the alignment.
5. Choose **Draw division** and click two endpoints spanning the property. Extend the endpoints to or beyond its edges. Repeat for additional divisions. **Undo division** removes the latest drawing operation.
6. Assign every numbered division to its corresponding Step 1 lot. Enter the plat reference and check the review box after verifying the lines and identities.
7. Choose **Apply to this map**, then use or download the map normally. Choose **Save divisions in calculation copy** to reuse the reviewed geometry when creating other map types or running V40's case-folder generation.

Verified GeoJSON polygons can also be imported. They must explicitly declare EPSG:2278 coordinates (US survey feet); each feature may provide a `legal_description` property for matching. Every division must have a unique Step 1 lot identity and cover the complete tax parcel without significant overlaps or gaps.

Internal platted lines appear dashed. Their visibility can be switched off independently of property outlines. Addresses are placed within the reviewed lot polygons. The editor's blue division labels and plat image are review aids and do not print. No automatic equal-strip geometry is generated from an `LTS` description, and map geometry does not overwrite verified Excel lot areas. A scan may require careful manual tracing; current replats must be considered when choosing the source.

## Council district and councilmember

Application import checks the selected properties against Houston's published council boundaries, then checks the district member in the City's official council directory. Both values populate Case Info. V40 also performs this lookup during case-folder preparation when the council option is selected.

For an existing workbook, choose **Find District / Councilmember**, review the result, and save a new calculation copy. Names remain editable in Case Info. If an application crosses district boundaries, the lookup returns the applicable districts and members. If the directory cannot be reached, a dated cached roster may be offered with a review warning; an unknown member is not invented.

The original calculation sheet is preserved when a new reviewed copy is saved. Prior map attachments remain available for review.
