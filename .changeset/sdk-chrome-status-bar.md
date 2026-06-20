---
"@casualoffice/sheets": minor
---

Chrome: add a minimal status bar to `<CasualSheets chrome="minimal" | "full">`.

Sits below the grid: Excel-style selection aggregates (Average / Count / Sum)
over the numeric cells in the active multi-cell selection, live. Self-contained
(reads the selection via `CasualSheetsAPI`). The richer status bar (configurable
stats, min/max, zoom, sheet tabs) lifts behind `"full"`.
