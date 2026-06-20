---
"@casualoffice/sheets": minor
---

Chrome: add a minimal formula bar to `<CasualSheets chrome="minimal" | "full">`.

Sits below the toolbar: a name box showing the active cell's A1 reference (live,
tracks selection) and an editable input showing its formula or value. Editing
commits through the facade — `=…` as a formula, numbers as numbers, else text.
Self-contained (reads the active cell via `CasualSheetsAPI`, no app context, no
autocomplete/name-box-dropdown/insert-function yet — those arrive when the rich
`apps/web` formula bar is lifted behind `"full"`).
