---
"@casualoffice/sheets": minor
---

`<CasualSheets onChange>` — a debounced stream of `IWorkbookData` snapshots.

The "host persists it" half of the Excalidraw model: the editor stays
storage-unaware and the host writes each snapshot wherever it likes
(localStorage, server, …). Driven by Univer's mutation hook
(`onMutationExecutedForCollab`), not UI events, so it captures every edit
including programmatic ones. Debounce window is configurable via
`onChangeDebounceMs` (default 400). Subscribed after the unit is created so
the initial mount mutations don't emit a spurious first snapshot.
