---
'@casualoffice/sheets': minor
---

Preserve data-bar conditional formatting through the xlsx round-trip, with the bar fill colour. ExcelJS can't carry a data bar faithfully — it reads everything except the fill colour and writes a broken `<color auto="1"/>` — so this adds a raw-OOXML bridge (`databar-passthrough.ts`, mirroring the pivot passthrough): the positive bar colour is read straight from the worksheet XML on import and the whole `<cfRule type="dataBar">` block is spliced into the worksheet XML on export. Imported data bars now render in-editor (via Univer's IDataBar) and round-trip their colour, min/max anchors, and show-value flag.

Scope is the legacy data-bar block (positive colour + min/max + showValue). The x14 extension — explicit gradient flag and negative/axis colours — is deferred; Excel renders a sensible gradient bar from the legacy block, and axis/border/direction have no representation in Univer's model. A data bar anchored to a `formula` threshold is dropped (ExcelJS floatifies cfvo values on read), consistent with the other CF rule types.
