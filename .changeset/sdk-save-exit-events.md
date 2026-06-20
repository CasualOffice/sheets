---
"@casualoffice/sheets": minor
---

`<CasualSheets>` save/exit events — the host-owned persistence contract (Phase 2).

- **`onSave(snapshot)`** — fired on Ctrl/Cmd+S inside the editor (the browser save
  dialog is suppressed). The host persists the snapshot.
- **`onExit(snapshot)`** — fired once on unmount with the final snapshot — the
  host's last chance to persist before the workbook is disposed.

With the existing `onChange`, these complete the "the SDK emits, the host stores —
never localStorage" model. The SDK still writes no storage of its own.
