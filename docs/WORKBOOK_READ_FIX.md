# v1.28: workbooks open in Excel or OneDrive

All four wizards now allow Windows file-replacement sharing while reading saved workbooks. This resolves a Permission denied error when Excel or a sync client holds a replacement handle but still permits reading.

Where an ordinary file open fails, the wizard reads through a Windows handle that shares read, write and replacement access. Calculation readers receive a private temporary copy. This never changes the original file, closes Excel, or changes file permissions.

The map's source fingerprint, preview loading, refresh, final save and later attachment validation use the same compatible reader. The workbook must still match the saved version used by the preview. Unsaved Excel edits are not included; save them before refreshing the map.

A short retry handles temporary exclusive saves. If another program denies every reader, the wizard explains how to finish saving and retry. An unreadable workbook cannot silently reuse old cached data.
