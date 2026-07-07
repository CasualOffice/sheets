---
'@casualoffice/sheets': minor
---

SDK parity with the unified docs/sheets contract (doc 38). `CasualSheetsAPI` gains
`undo()` / `redo()` (dispatch Univer's `univer.command.undo` / `.redo`, the same
command path the built-in chrome uses). New `mountCasualSheets(container, options)`
imperative entry mounts `<CasualSheets>` into a DOM node for non-React hosts and
resolves the full `CasualSheetsAPI` plus a `destroy()` — the sheets peer of the
docs SDK's `renderAsync`, distinct from the iframe-only `mountEmbedded`.
`AttachCollabOptions` field names are documented as aligned with the unified
`CollabConfig` shape (`server` / `room` / `password` / `token` / `role` / `share`).
Additive only — no existing export changed or removed.
