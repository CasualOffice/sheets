---
"@casualoffice/sheets": minor
---

`<CasualSheets chrome="none" | "minimal" | "full">` — the chrome scaffold.

First slice of the Office-chrome lift (SDK_MIGRATION_PIPELINE Phase 1 step 2).
`chrome="none"` (default) keeps the bare grid. `"minimal"` / `"full"` wrap the
grid in a flex column with a built-in toolbar (undo / redo / bold / italic /
underline) that drives the editor through `CasualSheetsAPI.executeCommand` — no
app context, no font dependency, works in any host. The rich Office shell
(formula bar, menus, status bar) is lifted from the app behind `"full"` in later
slices; until then `"minimal"` and `"full"` render the same toolbar.
