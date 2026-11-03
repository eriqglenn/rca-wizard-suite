# Property lot size in MLS notifications

MLS and MLS Renewal notification templates use the dedicated field `«Property_Lot_Size»` in this sentence:

> Preliminary staff review of the application indicates that your property's lot size is «Property_Lot_Size» square feet.

The value comes from that property's Step 1 land area. It is independent of the proposed or renewed minimum. The restriction paragraph continues to use the case minimum from the calculation sheet.

A separate letter is generated for each property, including properties sharing owner and mailing details. Numeric values use comma separators and retain any decimal portion. A missing or invalid property area displays `[VERIFY PROPERTY LOT SIZE]`; it never defaults to the case minimum.

Older copies of the MLS notification templates containing the previous preliminary-review sentence are supported automatically. Both SFR and No-SFR MLS Renewal notifications use the new field. MBL and MBL Renewal are unchanged by this lot-size correction.

Updated in place in v1.26 and v1.30 on September 9, 2026.

