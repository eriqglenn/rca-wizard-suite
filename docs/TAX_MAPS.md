# HCAD tax maps

All four wizards have Create Tax Map in Required setup. It reads the saved
calculation sheet and uses HCAD Parcel Viewer's parcel basemap and Letter
Landscape Side print template. One continuous yellow outer outline encloses all worksheet
properties, including excluded and unaddressed properties. It bridges streets and
gaps, preserving stepped parcel edges where possible, with no interior holes. HCAD's own lot and
block labels remain on the map.

Use More area, Closer, and the direction buttons to adjust the framing, then
Update Preview. The boundary always stays contiguous. Reset
framing restores north-up and the original extent. Refresh HCAD data refreshes
both the matched parcel data and the HCAD print.

Download Tax Map saves a PDF or PNG wherever you choose. Use This Tax Map saves
the reviewed PDF in the output folder and selects it in Tax map (required),
ready for the RCA packet. It does not replace the City boundary-map attachment.
Existing tax maps are retained; subsequent saves get unique filenames. Warnings
do not prevent downloading or using a completed preview.

Internet access to HCAD's public print service is required for a new view.
Identical prints are cached for the day. No ArcGIS account or GIS installation
is required. Only public map layers and boundary geometry are sent to HCAD's
print service; applicant and owner mailing information is not sent.

Developer checks: RCA Wizard Tools.exe taxmapcheck --out <audit folder>
--data <parcel_data.json> --live verifies the live HCAD print. --ui exercises
Create Tax Map, framing, download, and attachment selection in all four wizards.

--cases <cases.json> verifies completed workbooks for each wizard type.

Sources: https://arcweb.hcad.org/parcel-viewer-v2.0/ and its HCADMaps public
print service. Export payload follows Esri's ExportWebMap specification.

Only the outer boundary is emphasized. Interior parcels retain ordinary lines
and labels, without individual lot highlighting or extra bold interior rings.
