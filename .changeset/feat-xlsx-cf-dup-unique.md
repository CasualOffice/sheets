---
'@casualoffice/sheets': minor
---

Preserve duplicate-values and unique-values conditional formatting through the xlsx round-trip. ExcelJS drops these rules entirely (no reader or writer for them), so they're bridged via raw OOXML (`cf-dxf-passthrough.ts`): on import the `<cfRule type="duplicateValues|uniqueValues">` is read from the worksheet XML and its style resolved against `styles.xml`'s `<dxfs>`; on export the rule is spliced back into the worksheet XML and its differential style appended to `<dxfs>` (with correct index coordination against any styles ExcelJS already wrote). Imported duplicate/unique rules now render in-editor and round-trip their fill/font style.
