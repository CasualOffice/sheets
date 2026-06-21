---
'@casualoffice/sheets': minor
---

Config + events are now wired end-to-end through the iframe embed (not just the React component):

- **Feature flags over the wire** — hosts call `EmbedHostTransport.sendSetFeatures({ features })` (or `casual.command.set.features`) to disable any toolbar group / menu item / capability; the embed forwards it to the chrome (`features` prop), which hides the control and blocks its command.
- **Host-owned dialogs** — when a chrome control backed by a dialog the SDK doesn't render (Format Cells, Insert Chart, Find & Replace, …) is activated, the embed emits `casual.dialog.request`; hosts handle it via `EmbedHostTransport.on({ onDialogRequest })` and render their OWN dialog, applying the result via `executeCommand`. The React `<CasualSheets>` component exposes the same `features` + `onDialogRequest` props directly.

New exports: `CommandSetFeaturesData`, `DialogRequestData`, `ChromeTopProps`.
