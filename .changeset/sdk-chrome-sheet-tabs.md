---
'@casualoffice/sheets': minor
---

feat(chrome): sheet tabs in the built-in chrome

`<CasualSheets chrome>` now renders a worksheet tab strip above the status bar:
switch sheets (click), add a sheet (+), rename (double-click), and delete
(right-click → Delete, with the last visible sheet protected). Driven entirely
through the FUniver facade and kept live via the sheet-lifecycle events (plus a
mutation-level fallback so collab/replay-driven changes refresh too). Closes the
most fundamental gap between the SDK chrome and a real multi-sheet editor.
