---
'@casualoffice/sheets': minor
---

feat(chrome): custom Find & Replace dialog

`<CasualSheets chrome>` now has Find & Replace, opened with Ctrl/Cmd+F (find) or
Ctrl/Cmd+H (replace): match count + next/prev navigation (Enter / Shift+Enter),
match-case toggle, Replace / Replace All. It's a custom, facade-driven dialog —
search reads the active sheet's cells from `getSnapshot()`, navigation activates
the matching cell, replace writes via `setValue` — because Univer's own
find-replace UI doesn't render in the SDK's headless mount. Closes the last core
chrome gap.
