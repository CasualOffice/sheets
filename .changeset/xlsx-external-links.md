---
'@casualoffice/sheets': patch
---

Preserve external-workbook links across an xlsx round-trip. A formula like
`=[1]Sheet1!A1` references another workbook via `<externalReferences>` →
`xl/externalLinks/**`. ExcelJS has no external-link model, so it rebuilt the
export without those parts and without `<externalReferences>` — the `[N]` index
dangled and the formula resolved to `#REF!` on save (silent corruption). A new
external-link passthrough captures `xl/externalLinks/**` in reference order at
parse time and re-injects them at export, patching `[Content_Types].xml`,
re-creating the workbook→externalLink relationships, and rebuilding
`<externalReferences>` in the original order so the `[N]` indices still resolve.
The link parts (source path + cached values) are preserved verbatim.
