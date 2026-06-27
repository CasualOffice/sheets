---
'@sheet/web': patch
---

Make the keyboard-shortcuts cheat sheet accurate and more complete for Excel users. Four entries were wrong: `Ctrl++` / `Ctrl+-` were labelled "Zoom in/out" but actually insert/delete cells, `Shift+F11` was "Toggle full screen" but inserts a sheet, and `Ctrl+Shift+D` was "Refresh data" but shows pivot details. Corrected those and added the working-but-undocumented shortcuts Excel users reach for — Fill down/right, Copy from cell above, Insert/Delete cells, Go To, Find & Replace, AutoSum, Insert function, Recalculate, Toggle filter, Trace precedents/dependents, Insert sheet/chart, outside border, and grow/shrink font — grouped into Essentials / Editing / Navigation & selection / Formatting / Formulas & data / Insert & sheets. Also fixes `formatShortcut` dropping the literal `+` key so `Ctrl++` renders as `⌘+` / `Ctrl++` instead of a bare modifier.
