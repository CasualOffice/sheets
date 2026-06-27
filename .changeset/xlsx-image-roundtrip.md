---
'@casualoffice/sheets': patch
---

Preserve embedded images and shapes across an xlsx round-trip. Univer has no
drawing model, so ExcelJS rebuilt the exported workbook without any picture —
opening an `.xlsx` and saving it silently dropped every embedded image. A new
drawing-passthrough layer captures `xl/media/**` + `xl/drawings/**` (and the
per-sheet drawing linkage, keyed by decoded sheet name) at parse time and
re-injects them at export, patching `[Content_Types].xml` and the sheet
`<drawing>` relationship. Images aren't rendered in the editor yet, but they now
survive open → save so Excel keeps them. Shapes/SmartArt ride along (same parts).
