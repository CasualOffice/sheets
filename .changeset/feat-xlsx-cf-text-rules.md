---
'@casualoffice/sheets': minor
---

Preserve more conditional-formatting rule types through the xlsx round-trip. On top of the existing `cellIs` (numeric) and `expression` (formula) highlight rules, the bridge now maps Excel's `top10` (top/bottom N, with percent), `aboveAverage` (above/below the range mean), `timePeriod` (today / last 7 days / this month / …), and `containsText` text rules (the `containsText` operator plus the blanks/errors predicates) to and from Univer's conditional-formatting resource — so these survive Excel → open here → save → Excel and paint immediately on open.

`beginsWith` / `endsWith` / `notContainsText`, `duplicateValues` / `uniqueValues`, and the visual rule types (color scale / data bar / icon set) remain unmapped: ExcelJS can't round-trip the first two groups without losing their meaning, so they're skipped rather than corrupted.
