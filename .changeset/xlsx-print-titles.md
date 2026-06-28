---
'@casualoffice/sheets': patch
---

Preserve print titles (repeat rows/columns) across an xlsx round-trip. The
page-setup bridge already carried orientation, scale, margins, and print area;
it now also captures `printTitlesRow` / `printTitlesColumn` on import and
re-applies them on export, so a workbook authored to repeat its header row (or
left column) on every printed page keeps that setting through open → save.
