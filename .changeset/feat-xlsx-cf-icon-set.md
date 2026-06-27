---
'@casualoffice/sheets': minor
---

Preserve icon-set conditional formatting through the xlsx round-trip. Excel icon-set rules (3/4/5 arrows, traffic lights, signs, ratings, flags, symbols, quarters) now map to and from Univer's conditional-formatting resource — the icon group, threshold bands, `reverse` flag, and show-value option round-trip, and the correct icon paints per band immediately on open.

OOXML orders icon thresholds low→high while Univer's bands run high→low (top icon first), so the mapping inverts threshold order on import and back on export. The three Excel-2010 x14 icon sets (`3Triangles` / `3Stars` / `5Boxes`) and any icon set using a `formula` threshold are skipped — ExcelJS can't write them faithfully, so they're dropped rather than corrupted.

Also re-applies a fix that didn't land in the previous CF merge: a color-scale rule using a `formula` threshold is now dropped (ExcelJS floatifies the threshold value on read, destroying the formula) instead of emitting a corrupt `NaN` stop.
