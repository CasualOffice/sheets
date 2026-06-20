---
'@casualoffice/sheets': minor
---

feat(chrome): AutoSum dropdown in the toolbar

The built-in chrome toolbar gains an Excel-style AutoSum control (Σ): Sum /
Average / Count numbers / Max / Min. Picking one inserts `=FN(<selection>)` one
row below a multi-cell selection (and activates that cell), or `=FN()` into a
single active cell. Pure facade — no Univer UI dependency — so it works in the
embedded mount.
