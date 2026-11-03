# Create Boundary Map - v1.26

1. Save your case calculation sheet in Excel.
2. In the correct wizard, select that workbook and the case output folder.
3. Click **Create Boundary Map** under Required setup.
4. Review the property count, description, minimum value, boundary and labels.
5. Adjust orientation or framing if needed. Drag labels to reposition them.
6. Click **Download Map** to choose a PDF or PNG file and save location. Downloading
   keeps the preview open and does not change the selected attachment.
7. Click **Use This Map** to save and select the PDF for the existing letter,
   ordinance and RCA attachment workflows.

The heading is Special Minimum Lot Size, Special Minimum Building Line, or the
corresponding title followed by Renewal. MLS values use Square Feet; MBL values
use Feet. The footer preserves the measured City Calibri fonts, spacing, boundary
key, source/date/reference and disclaimer. Street labels are all Arial 18 pt.
The lower-right legend contains the original supplied City of Houston seal and
Planning & Development Department wordmark, beneath the disclaimer. Its size is
132 by 52.8 points, preserving the artwork's proportions and matching the
reference placement. The asset ships in Wizard_Files/assets/maps for all four
wizards; no file from the user's Downloads folder is needed after installation.
Block ranges use a hyphen: `100 and 200 blocks` becomes `100-200 blocks`, and
`2400- 2500 block` becomes `2400-2500 block`. This applies only to numeric block
ranges; cross-street and side descriptions keep their normal wording.

## Using a map you already have

Select your preferred boundary map with Browse. If it was made with an older
map layout, a different wizard type or a calculation sheet that has since
changed, an amber warning appears beside Create Boundary Map. You can continue
generating letters, ordinances and RCA documents with that selected map.
Updating the map is optional; no confirmation or regeneration is required.
The wizard does not replace your chosen map or alter the calculation sheet.

## Data sources

Step 1 supplies the property list and accounts, including leading zeros.
Step 2 supplies the case number, street, block description, sides and cross
streets. Its property-table Land Use values take precedence over Step 1's mirrored
values. Supported classifications are SFR, MF, COM, VAC and EXC. Every property
listed in the calculation sheet is inside the application boundary, including
EXC parcels. Land-use labels and the land-use key are omitted from the map.
This does not change worksheet calculations or document-generation rules.

The account is checked against the worksheet address. When an account is stale,
the generator looks up the exact property address. A matching legal description
resolves multiple parcels at that address; unresolved ambiguity still stops the
map. The preview reports account changes, and the saved metadata records both
the worksheet account and the matched account. Workbooks are never rewritten.
Unaddressed parcels use their account and legal description for matching and
display **0** on the map. Zero indicates an unaddressed lot; it does not change
the source worksheet or supply a street number for parcel matching.
Numbered street names recognize an omitted ordinal suffix, such as E 22 Street
for E 22ND ST. Map street labels use the corresponding City centerline name.
If a Step 2 street is absent from the City layer, the program checks Harris
County's CTS street geometry and recorded aliases in the same area. A notice
identifies the alternate source. The description always keeps Step 2 wording.
If neither source locates a street, the map still opens and can be saved with
a warning. Its label is movable; review and drag it to the correct street.
Missing street services use north-up framing when the main street cannot be
located. Street-label placement warnings do not disable **Use This Map**.
Warnings and data sources are retained in the saved provenance file.

New MLS uses the saved Step 2 E4 result. New MBL uses L16, or its E4 label when
L16 is empty. The 2026 renewal sheets use D2; the recognized older renewal layout
uses E10. The generator uses the workbook recalculation helper, but never derives
a substitute minimum from individual property areas or building measurements.

Public HCAD parcel data:
https://www.gis.hctx.net/arcgis/rest/services/HCAD/Parcels/MapServer/0

Houston Map Viewer retained parcel fallback:
https://city-gis.internal.example/arcgis/rest/services/DIR/LandbaseandRoads_gx/MapServer/6

When an old account is missing from the county layer, the exact City parcel is
checked. If the City retains every matched account, one consistent City parcel
edition is used throughout the map. This avoids mixing retired parcels with
newer combined parcels. The preview reports this choice and the recorded tax year.

City street centerlines:
https://city-gis.internal.example/arcgis/rest/services/DIR/LandbaseandRoads_gx/MapServer/9

Queries request coordinates and necessary parcel/street attributes, not owner
mailing information. Object IDs are fetched first and geometry retrieved in
bounded batches, as supported by the [ArcGIS query API](https://developers.arcgis.com/rest/services-reference/enterprise/query-map-service-layer/).
Returned IDs, projection and completeness are checked. Coordinates and scales
use EPSG:2278 US survey feet. Outlines retain multipart parcels and holes.

## Preview and saved files

The default orientation fits the main parcel row and actual end intersections
on the page. Addresses are placed entirely inside their own matched parcels.
Standard address text is Arial 12 pt; narrow lots use automatic rotation and,
when needed, a smaller size down to 5 pt. Exceptionally small lots can use text
down to 2 pt with a warning. Fitting measures the actual Arial
glyph outlines. It never moves an address outside the lot or adds a callout.
Street labels remain Arial 18 pt. North up, rotation, zoom, pan and 90-degree
address-label rotation are available. Labels can be moved by dragging; an
address move outside its own parcel is rejected and reported in the preview.
The program checks for clipped parcels, clipped labels, overlapping labels and
address labels outside their parcel. All preview layout issues are warnings;
**Download Map** and **Use This Map** remain available once a PDF is created.
Addresses stay inside their own lots. If a parcel is too small to contain text,
its missing label is reported for review and does not prevent saving.
Case facts are edited in Excel, not separately in the map preview.

For two-sided cases, Enclose the street between parcel rows starts selected.
It closes street-width gaps while retaining actual parcel edges and subtracting
unlisted parcels. Listed EXC properties remain inside. Inspect this enclosure for the application's
intended boundary. Turn it off to outline parcel edges individually. This is
an application boundary map; MBL building setbacks are not measured or drawn.

Use This Map saves the exact reviewed PDF, a PNG and a `.boundary.json` record of
the input fingerprint, sheet rows/accounts, layout, scale and data sources. The
PDF also contains the input fingerprint so the workbook-change check still works
when it is copied without its sidecar. Existing files are never silently replaced.
Older maps can be recreated with Create Boundary Map to add the City seal.
The address-only layout keeps each address inside its lot and EXC properties enclosed.

Saved public data is reused from the output folder's `.boundary_map_cache`.
The preview shows original retrieval dates; Refresh parcel data bypasses that
cache. A retrieval date does not change the assessment year of HCAD's records.

## Cases requiring correction or another map

Resolve missing, ambiguous, overlapping or conflicting account/address matches
in the workbook. Unit information can differ from the public parcel's base
street address when the account and base address agree. Duplicate rows for the
same account, address and land use are drawn once and reported in the preview.

The installed default MBL sheet has fictitious heading streets and five real
Haver Street property rows. Its Create Boundary Map preview is explicitly marked
MBL EXAMPLE and uses those five north-side parcels between Windsor and Mandell.
This exception applies only to the unchanged example parcel set in the installed
instructional sheet. A planner's case workbook always supplies its own heading.
Other unrelated property-list headers still require correction. An unmatched
cross street gives a warning and allows map creation. The sample workbooks are
unchanged.
The generator handles full tax-parcel boundaries. It cannot establish a surveyed
partial-lot line or resolve overlapping/stacked parcel geometry automatically.
Very large cases, exceptionally thin parcels or long descriptions may not fit
a single page at readable sizes. Review the warnings and adjust framing or
orientation when needed. A generated map can still be downloaded or used with
these warnings. If the workbook changes during preview, the map keeps the
displayed case and records a warning. The existing Browse attachment option
remains available.

## Verification

Version 1.25 adds checks for alternate county street names and recorded aliases,
missing cross streets in all four wizards, failed street services, cancellation,
and saving with street-label warnings. Its 36 map tests pass. The MBL 93REN
Rosedale calculation sheet creates and attaches all 13 parcels using county
Hutchins Street geometry, with its Step 2 renewal description and 35-foot value.
The four established case/example workflows also pass; source sheets are unchanged.

`tools/test_boundary_map.py` checks all four workbook adapters, typography,
minimum-value handling, accounts, exclusions/holes, query batching, stale maps,
preview attachment and saving. `tools/test_workflow.py` checks that existing
workbook selection, recalculation and session behavior is preserved.

The live Banks MLS 926 check matches 28 unique parcels from 29 sheet rows and
renders the correct 6,250-square-foot map. Crystalwood's saved public shapes are
also exercised through all four adapters with explicit test workbooks and values.
Those QA fixtures are not official cases or changes to a user's calculation sheet.

Version 1.21 additionally exercises the actual Create Boundary Map and Use This
Map buttons with MLS 905 (22 parcels), the default MBL example (5), MLS Renewal
247 (29) and MBL Renewal 91 (38). Each must render, save and attach successfully,
with the source workbook fingerprint unchanged. `rca_tools.py mapcasecheck`
runs this workflow from a JSON list of workbook paths into a separate test folder.

Version 1.24 expands verification to 14 existing calculation sheets. Ten create
maps and attach their saved PDF; four correctly report conflicting worksheet
data. The successful sheets cover nine different streets and all four wizards.
The 32 regression tests include block ranges, numbered street names, old workbook
files without dimension metadata, North up, rotation, framing, panning, address
rotation, street enclosure, label dragging, reset, refresh, service-error retry,
closing, stale-map checks, and saved outputs. Closed Tk objects are finalized on
the main thread before another case's background workbook reader starts.

Ordinal suffixes in the legend description use lowercase (22nd, 18th, 21st, 23rd). Street labels inside the map keep their existing uppercase style.

Only the outer boundary is emphasized. Interior parcels retain ordinary lines
and labels, without individual lot highlighting or extra bold interior rings.
