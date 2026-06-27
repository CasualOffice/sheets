---
'@casualoffice/sheets': patch
---

Speed up editing on large sheets. The Univer fork's `updateFormulaData` (run on every cell edit) did two whole-sheet O(cells) scans per edit and rebuilt the formula data for every sheet of every unit — so a single edit on a 100k-row workbook took ~124ms. It's now incremental: the formula-id map and the id→formula fix-up pass run only when the edit actually touches a shared-formula (`si`) relationship, and it seeds from just the edited sheet. A single-cell edit on a 100k-row sheet of SUM formulas drops to ~77ms. Validated by the fork's full formula test suite (3834 tests) plus end-to-end recalc checks.
