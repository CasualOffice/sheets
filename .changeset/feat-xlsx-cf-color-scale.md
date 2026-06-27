---
'@casualoffice/sheets': minor
---

Preserve color-scale conditional formatting through the xlsx round-trip. Excel's 2- and 3-color scales now map to and from Univer's conditional-formatting resource — the gradient stops (`min` / `max` / `num` / `percent` / `percentile` / `formula` thresholds, each with its color) round-trip and paint the value-mapped gradient immediately on open.

Also re-applies a hardening fix that didn't land in the previous CF merge: a foreign or partially-formed CF resource payload with no `style` no longer throws and aborts the whole xlsx export.

`dataBar` remains unmapped — ExcelJS surfaces data bars via the x14 extension on read without the fill color, so they can't round-trip faithfully yet; `iconSet` is pending (needs the OOXML↔Univer icon-ordering mapping). The text `beginsWith` / `endsWith` / `notContainsText` operators and `duplicateValues` / `uniqueValues` stay unmapped (ExcelJS can't round-trip them).
