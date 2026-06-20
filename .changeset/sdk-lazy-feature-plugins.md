---
"@casualoffice/sheets": minor
---

`<CasualSheets>` now lazy-loads the feature plugins by default (`lazyPlugins`,
default `true`): conditional formatting, data validation, hyperlinks, notes,
tables, comments, drawings, sort, filter, and find/replace.

Plugins whose data already lives in `initialData` (CF rules, tables, hyperlinks,
…) load eagerly *before* the workbook mounts, so opening a file never silently
drops them; everything else idle-loads after first paint. This brings the SDK
editor to feature parity with the app's grid without bloating the initial
chunk — `@univerjs` feature packages stay external and load on demand.

Pass `lazyPlugins={false}` for the minimal editor (render + formula + numfmt
only); the embed-iframe runtime sets this to remain a single self-contained
bundle.
