---
'@casualoffice/sheets': minor
---

Add `CasualSheetsAPI.exportXlsx()` and a `workbookDataToXlsx` converter on the `@casualoffice/sheets/xlsx` subpath — the SDK is now a two-way xlsx I/O surface (was import-only). The core converter (values/formulas, styles, merges, number formats, borders, hyperlinks, comments, data validation, tables, page setup, named ranges, VBA passthrough) was lifted out of `apps/web` and runs in its own Web Worker; ExcelJS stays out of the editor entry (lazy-loaded as a separate chunk). App-level feature models (charts/pivots/sparklines) remain a power-host concern, baked into the snapshot before serialization via the generic `ExportExtras` (`hyperlinks` / `outline` / `chartImages`).
