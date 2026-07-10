# @sheet/web

## 0.5.2

### Patch Changes

- Updated dependencies [6a9c647]
  - @casualoffice/sheets@0.18.0

## 0.5.1

### Patch Changes

- Updated dependencies [17b3a23]
  - @casualoffice/sheets@0.17.0

## 0.5.0

### Minor Changes

- 5bc2ecc: Add **named cell styles** (Format ▸ Cell styles): Excel's Good / Bad / Neutral semantic styles + Title / Heading 1 / Heading 2 + Normal (reset) — composite fill/font-colour/bold/size applied to the selection. Closes a gap vs OnlyOffice + Google Sheets.
- 85850f7: Add **combo charts + a secondary (dual) value axis** (chart Format dialog ▸ "Series type & axis"). Per-series controls let a column / line / area chart mix bars and lines (Excel's Combo chart type) and plot any series against a secondary right-hand value axis — the chart then renders two `yAxis` entries with the flagged series routed to `yAxisIndex: 1`. Overrides persist on the chart's `format` (so they survive reload + xlsx round-trip), default to the chart's base type/axis, and are gated to the families where they read correctly (not pie / doughnut / scatter / 100%-stacked / horizontal bar). Closes a common gap vs Excel for revenue-vs-margin style charts.
- 41ce9ea: Add **Download as PDF** (File menu). Exports the active sheet's used range as a real, searchable vector PDF (jsPDF + jspdf-autotable) — paginated across A4 pages, with column letters, row numbers, a workbook/sheet title, and **per-cell styling carried over** (bold, horizontal alignment, fill colour). Closes a MUST-have gap vs OnlyOffice + Google Sheets (we previously only had Print → "Save as PDF").
- df79443: Add **PivotTable column fields** (cross-tab / matrix layout). The Insert ▸ PivotTable dialog gains a **Column field** dropdown: placing a field on the column axis fans the value field out across one column per distinct value, producing Excel's matrix layout (`Region` down the rows, `Quarter` across the columns) with a right-hand **Grand Total** column carrying row totals, a bottom Grand Total row carrying column totals, and the overall total in the corner. Works with multi-row compact layouts and filters; drill-down ("show details") now narrows by the clicked column's key as well as the row key. Picking the same field for both rows and columns falls back to the classic row-only pivot. Closes the column-field gap noted in the P0 pivot spec.
- a548717: Add **Protect (read-only)** toggle (Data menu). Locks the workbook so edits are blocked, reusing the SDK's `applyReadOnly` engine (command veto + permission flip); toggling again lifts protection. Closes a gap vs OnlyOffice + Google Sheets (sheet/range protection).
- 2ff6454: Add **text rotation** (Format Cells ▸ Alignment ▸ Text orientation): None / ±45° / ±90°, applied via Univer's `FRange.setTextRotation`. Closes a gap vs OnlyOffice + Google Sheets.
- c32a5f5: Real-time comment sync across co-editors. Thread comments (add / edit / resolve / delete / re-anchor) now cross the Yjs op-log bridge so a comment created, resolved, or deleted by one collaborator shows up live for everyone in the room — previously they stayed local-only. The five `thread-comment.mutation.*` ids are added to the bridge's `SYNCED_MUTATIONS` allowlist and mapped to the `threadComment` lazy-plugin group so a peer that hasn't opened the Comments pane loads the plugin before replaying the change (no silent drops). The `threadComment` lazy loader now also pulls `@univerjs/sheets-thread-comment/facade`, installing `FWorksheet.getComments` / `FRange.addComment` so the Comments pane and facade-driven flows work once the plugin mounts. Covered by a new two-client `coedit-comments` Playwright spec asserting bidirectional propagation.
- 94131b8: Data Validation parity: add the **Time** Allow-type and fix DV i18n.
  - **Time validation type** — Excel exposes Time as a distinct Allow-type (Whole / Decimal / List / Date / **Time** / Text length / Custom). Univer's `DataValidationType.TIME` enum and the cell-edit time-picker already existed but no validator/view was registered; the fork now registers `TimeValidator` (parses to a fractional serial, validates the standard operators, normalizes to `HH:mm:ss`) and its panel view.
  - **Input Message editor** — the DV panel's Advance options now expose the input-message toggle + title/text fields (the on-hover popup shipped previously).
  - **i18n fix** — the locale bundle merged only the DV _UI_ strings, so the DV Type/Operator selectors and cell error messages rendered raw i18n keys (`sheets-data-validation.date.title` instead of "Date") for every type. The base `@univerjs/data-validation` + `@univerjs/sheets-data-validation` locales are now merged in both the app and the SDK embed runtime.

- fd1ea16: Native pivot export (opt-in). The SDK now exports `generateNativePivot` (build real `xl/pivotTables` + `xl/pivotCaches` OOXML from a pivot model) and `applyPivotsToZip`, so a host can compose native PivotTables into an export. The app wires this behind an off-by-default flag (`cs-native-pivots`): when enabled, in-app pivots round-trip to Excel as real, refreshable PivotTables instead of flat cells. Default behaviour is unchanged.

### Patch Changes

- 650d803: Fix multi-hundred-millisecond lag when selecting a whole column or row. The status-bar Sum/Avg/Count recompute scanned the selection's full nominal extent (up to the sheet's 1,048,576 rows) synchronously on the main thread; it now bounds the scan to the used range (last row/column with content), which is exact for all statistics and matches how Excel computes column stats instantly. Fixed in the Univer fork (`status-bar.controller.ts`).
- 3afcc8d: Harden the desktop save bridge against two data-loss vectors: reject an empty (0-byte) serialization in `chunkedWrite` instead of atomically committing it over the original file, and stop clearing the dirty flag after a save if an edit landed while the write was in flight (it would otherwise be marked saved and lost on window close).
- 5961120: Replace the desktop runtime-error overlay's full-width red monospace banner (which read like a crash/dev artifact across the grid) with a compact, dismissible notice; full error detail goes to the console.
- 0d27d09: Fail clearly when a desktop file is truncated or replaced mid-open. The chunked read sizes its buffer to `document_size`, so if the file shrinks between sizing and reading (another process truncates or replaces it — the same external edits the file watcher reports), the tail was left zero-padded and parsed as a baffling "corrupt spreadsheet". The bridge now detects the short read and throws a clear "the file changed while opening — try again" error instead of returning a silently-mangled buffer.
- a069d89: Retry the desktop external-change reload once on a transient failure. When another app saves the open file, the filesystem watcher often fires while the write is still in flight (an atomic save briefly truncates/replaces the file), so the first reload reads short and threw — previously swallowed, leaving the user on stale content. The reload now retries once after a short settle delay, by which point the external write has completed and the reload succeeds.
- 4930f74: Close a narrow data-loss window in the desktop save path: the bridge pinned its "did the doc change" reference at `bridge.save()` entry, but the bytes are serialized earlier (chart render + xlsx encode happen in between). An edit landing in that gap could leave the on-disk file stale while the window was marked clean — and silently lost on close. The save caller now captures the edit counter at serialization time (`wb.save()`) and passes it through as `save(bytes, baselineSeq)`, so the bridge only clears dirty when nothing changed since the bytes were produced.
- c6e4ed5: Serialize desktop file writes so two overlapping saves can't corrupt the file. A fast double Ctrl+S — or a Ctrl+S issued while a large save is still streaming chunks — previously ran two `begin_save_document` / `write_save_chunk` / `commit_save_document` sequences concurrently against the same per-path temp file, which could interleave and produce a corrupt or truncated result. Writes now run through a chain so each save fully completes before the next begins; a failed save no longer wedges later ones, and its error still surfaces to the caller.
- cc1eb26: Make the keyboard-shortcuts cheat sheet accurate and more complete for Excel users. Four entries were wrong: `Ctrl++` / `Ctrl+-` were labelled "Zoom in/out" but actually insert/delete cells, `Shift+F11` was "Toggle full screen" but inserts a sheet, and `Ctrl+Shift+D` was "Refresh data" but shows pivot details. Corrected those and added the working-but-undocumented shortcuts Excel users reach for — Fill down/right, Copy from cell above, Insert/Delete cells, Go To, Find & Replace, AutoSum, Insert function, Recalculate, Toggle filter, Trace precedents/dependents, Insert sheet/chart, outside border, and grow/shrink font — grouped into Essentials / Editing / Navigation & selection / Formatting / Formulas & data / Insert & sheets. Also fixes `formatShortcut` dropping the literal `+` key so `Ctrl++` renders as `⌘+` / `Ctrl++` instead of a bare modifier.
- 2c0cb1f: Fix a desktop open race where a workbook that parsed before Univer finished booting was dropped, leaving a blank grid still bound to the real file (a subsequent edit + save could then overwrite the original .xlsx with an empty workbook). The revision swap effect no longer advances its revision marker when the api isn't ready yet; it waits and re-applies the pending snapshot once the editor signals ready.
- 266082b: Fix File → New overwriting the previously-open file in the desktop app. New replaced the workbook in-window but left the bridge's bound file path pointing at the old file, so the next Save wrote the empty workbook over it. New now clears the bound path so Save prompts for a location.
- c990c9f: Fix File → Open overwriting the previously-open file in the desktop app. Opening a file in-window replaced the workbook but left the bridge bound to the old file path, so the next Save wrote the newly-opened content over it. Open now unbinds the path so Save prompts for a location.
- 6a2331c: Fix raw i18n keys in the filter, table and hyperlink features. Like the data-validation fix in #252, the locale bundle merged only the `-ui` halves of these features, so their error toasts and generated labels rendered raw keys (e.g. `sheets-table.tablePrefix`, `sheets-filter.command.not-valid-filter-range`, `sheets-hyper-link.message.refError`). The base `@univerjs/sheets-filter`, `@univerjs/sheets-table` and `@univerjs/sheets-hyper-link` locales are now merged into both the app bundle and the SDK embed runtime. (All three are already pinned + fork-linked deps, so no dependency changes.)
- 7058444: Stop autosave from hitching the grid on large workbooks. The autosave snapshot is a full `wb.save()` deep clone that can take hundreds of milliseconds on a big sheet, and the 30-second tick ran it on the main thread regardless of what the user was doing — freezing the grid mid-keystroke. The snapshot now runs in a `requestIdleCallback` slot (with a 2s timeout guarantee, and a `setTimeout` fallback for older Safari), so it lands when the browser is idle rather than mid-edit. The `pagehide`/`beforeunload` flush stays synchronous, and the saved content is unchanged — only the timing moves.
- 8d81712: Defer the version-history auto-snapshot to an idle slot too, matching the autosave fix. Its ~10-minute capture interval ran a full `wb.save()` deep clone on the main thread regardless of activity, which could freeze the grid mid-edit on a large workbook. The shared `runWhenIdle` helper (extracted to `idle.ts`, now used by both the autosave and version-history loops) runs the clone when the browser is idle instead.
- dc68c7d: Avoid a synchronous full-range permission scan when dragging the fill handle from a very large (e.g. whole-column) selection on a sheet with no protection rules. The fill-handle permission check now short-circuits when the fill range overlaps no protected range — matching the range-move check — instead of walking every cell. Fixed in the Univer fork.
- Updated dependencies [94131b8]
- Updated dependencies [6936318]
- Updated dependencies [6dff888]
- Updated dependencies [1411f01]
- Updated dependencies [962161f]
- Updated dependencies [40c12bd]
- Updated dependencies [4720ed5]
- Updated dependencies [0d508b5]
- Updated dependencies [fd53802]
- Updated dependencies [d380a3e]
- Updated dependencies [6cd6417]
- Updated dependencies [1651f35]
- Updated dependencies [680715c]
- Updated dependencies [86268d7]
- Updated dependencies [6a2331c]
- Updated dependencies [fd1ea16]
- Updated dependencies [5386225]
- Updated dependencies [5317557]
- Updated dependencies [fe2d4c9]
- Updated dependencies [251481e]
- Updated dependencies [c5911ae]
- Updated dependencies [41465e3]
- Updated dependencies [2306f83]
- Updated dependencies [0290b09]
- Updated dependencies [a7fa701]
  - @casualoffice/sheets@0.16.0

## 0.4.7

### Patch Changes

- Updated dependencies [7606355]
  - @casualoffice/sheets@0.15.1

## 0.4.6

### Patch Changes

- Updated dependencies [8bb728f]
- Updated dependencies [a3493fb]
- Updated dependencies [a921179]
  - @casualoffice/sheets@0.15.0

## 0.4.5

### Patch Changes

- Updated dependencies [224fc2c]
- Updated dependencies [33ded85]
- Updated dependencies [dfc8e6b]
  - @casualoffice/sheets@0.14.0

## 0.4.4

### Patch Changes

- Updated dependencies [2846392]
  - @casualoffice/sheets@0.13.0

## 0.4.3

### Patch Changes

- Updated dependencies [8b35360]
- Updated dependencies [58ce6a0]
- Updated dependencies [1adc983]
  - @casualoffice/sheets@0.12.0

## 0.4.2

### Patch Changes

- Updated dependencies [971ad7d]
  - @casualoffice/sheets@0.11.1

## 0.4.1

### Patch Changes

- Updated dependencies [6c8a94e]
- Updated dependencies [3c93042]
  - @casualoffice/sheets@0.11.0

## 0.4.0

### Minor Changes

- 4fd30c5: apps/web shares the SDK editor core (Phase 3 step 1)

  `apps/web` no longer hand-rolls its Univer bootstrap — `UniverSheet.tsx` now
  renders `<CasualSheets chrome="none">` from `@casualoffice/sheets`, sharing the
  SDK's Univer boot, plugin set, formula engine, and snapshot/API. The app keeps
  its rich shell (ribbon, charts, pivots, panels, dialogs) and layers its extras
  on top: crosshair-highlight + zen-editor + Merge/Unmerge context menu via
  `onBeforeCreateUnit`, off-main compute via `formula={{ worker }}`, and the
  paste-merge hook / dev helpers / zoom-shortcut override via `onReady`. One Univer
  bootstrap now serves both the app and third-party SDK hosts.

### Patch Changes

- Updated dependencies [49a3215]
- Updated dependencies [5256f3d]
- Updated dependencies [7f42243]
- Updated dependencies [29744e8]
- Updated dependencies [ce87187]
- Updated dependencies [99b617f]
- Updated dependencies [f6b1b24]
- Updated dependencies [67e0d55]
- Updated dependencies [7816a5d]
- Updated dependencies [1495444]
- Updated dependencies [838ce1b]
- Updated dependencies [3d9d0b5]
- Updated dependencies [35abbab]
- Updated dependencies [91ff777]
- Updated dependencies [f8b05b4]
- Updated dependencies [a090e65]
- Updated dependencies [65124b4]
- Updated dependencies [53b87fe]
- Updated dependencies [ea014be]
- Updated dependencies [f0d5779]
- Updated dependencies [c007f64]
- Updated dependencies [161aa91]
- Updated dependencies [3c5a990]
  - @casualoffice/sheets@0.10.0

## 0.3.13

### Patch Changes

- Updated dependencies [652068f]
- Updated dependencies [f93fa6c]
- Updated dependencies [d3f9be6]
- Updated dependencies [1da029e]
- Updated dependencies [2381fb4]
  - @casualoffice/sheets@0.9.0

## 0.3.12

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.8.0

## 0.3.11

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.7.0

## 0.3.10

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.6.0

## 0.3.9

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.5.7

## 0.3.8

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.5.6

## 0.3.7

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.5.5

## 0.3.6

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.5.4

## 0.3.5

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.5.3

## 0.3.4

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.5.2

## 0.3.3

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.5.1

## 0.3.2

### Patch Changes

- Updated dependencies [e044efd]
  - @casualoffice/sheets@0.5.0

## 0.3.1

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.4.0
