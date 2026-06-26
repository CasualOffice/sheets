---
'@casualoffice/sheets': patch
---

Fix xlsx import corrupting filled-down (shared) formulas. A slave cell of a shared formula was imported as `=<masterAddress>` (e.g. `=B1`) instead of its position-translated formula, so opening a workbook with an autofilled column and recalculating/saving silently corrupted every cell after the first. The parser now reads ExcelJS's translated `cell.formula` for shared-formula cells.
