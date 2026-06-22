# @casualoffice/sheets

## 0.15.0

### Minor Changes

- 8bb728f: Config + events are now wired end-to-end through the iframe embed (not just the React component):
  - **Feature flags over the wire** — hosts call `EmbedHostTransport.sendSetFeatures({ features })` (or `casual.command.set.features`) to disable any toolbar group / menu item / capability; the embed forwards it to the chrome (`features` prop), which hides the control and blocks its command.
  - **Host-owned dialogs** — when a chrome control backed by a dialog the SDK doesn't render (Format Cells, Insert Chart, Find & Replace, …) is activated, the embed emits `casual.dialog.request`; hosts handle it via `EmbedHostTransport.on({ onDialogRequest })` and render their OWN dialog, applying the result via `executeCommand`. The React `<CasualSheets>` component exposes the same `features` + `onDialogRequest` props directly.

  New exports: `CommandSetFeaturesData`, `DialogRequestData`, `ChromeTopProps`.

- a3493fb: The SDK chrome now has **built-in dialogs** and a **host extension API** — the two things that made it a viewer rather than a real editor.
  - **Dialog host**: chrome controls open a built-in dialog by default. Ships **Format Cells** (Number/Alignment/Font/Border/Fill) and wires **Find & Replace** into the dialog host. A `Dialog` primitive (portal + focus-trap + Escape/backdrop) backs them.
  - **Extension API** (`extensions` prop on `<CasualSheets>`, also forwarded through the iframe chrome): hosts add their **own** toolbar buttons, menu items, side panels, and dialogs on top of the built-ins — or **override** a built-in dialog by kind. Resolution order: host React override (`extensions.dialogs[kind]`) → host-owned via `onDialogRequest`/`hostOwnedDialogs` → SDK built-in. `onDialogRequest` stays backward-compatible. New exports: `ChromeExtensions`, `ToolbarExtension`, `MenuExtension`, `PanelExtension`, `DialogExtension`, `DialogComponentProps`, `PanelComponentProps`, `DialogKind`.
  - **Fix**: the sheet-tab right-click menu rendered behind the bars below the tab strip (stacking-context trap) — it's now portaled to `<body>` and opens upward, fully visible.

  Remaining built-in dialogs (Goal Seek, Name Manager, Page Setup, Paste Special, Insert Cells, Properties, Keyboard Shortcuts, About, Watermark, Insert Chart/PivotTable) and the side-panel rail land next.

### Patch Changes

- a921179: Two embed fixes found opening real files in a host (Drive):
  - **Page hung on open.** The embed ran formula recalc on the MAIN thread (no worker), so a formula-heavy workbook froze the page. The embed now bundles a formula worker (`formula.worker.js`, mirroring the reference app) and wires it via `<CasualSheets formula={{ worker }}>`, so compute runs off-thread. Verified: `=1+2*3` → 7 across 300 formula cells with the page staying responsive.
  - **Feature-plugin UIs showed raw i18n keys** (the comment panel rendered `thread-comment-ui.editor.reply` instead of "Reply"; tables/filter/CF dialogs likewise). The embed locale bundle only had the base plugins; it now includes every lazily-loaded feature plugin's en-US strings (comments, tables, sort, filter, conditional formatting, data validation, drawing, hyperlinks, notes, find/replace).

## 0.14.0

### Minor Changes

- 224fc2c: SDK chrome (`chrome="full"` / the iframe embed) now matches the real app's Home-tab toolbar + menus. **Toolbar** gains: font family/size selectors + grow/shrink, clipboard (paste/cut/copy/paste-values), format painter, text & fill color pickers, borders, vertical align + wrap text, a number-format dropdown, and AutoSum. **Menus** gain the full Edit/View/Insert/Format/Data/Help sets (freeze panes, show formulas, gridlines, insert sheet/table/image/hyperlink/comment, number-format submenu, increase/decrease decimals, borders, sort/filter/recalculate, etc.) — all driven purely through the FUniver facade + the same Univer command ids the app uses.

  Two new optional props on the chrome:
  - `features?: Record<string, boolean>` — hide any control/group (and block its command) when its flag is false. Lets hosts disable features.
  - `onDialogRequest?: (kind, context?) => void` — controls backed by a dialog the SDK doesn't ship yet (Format Cells, Insert Chart, PivotTable, Find & Replace, …) call this so the host can render its OWN dialog; without it they're omitted (no fake dialog). Built-in dialogs land in a later release.

- 33ded85: The iframe embed now ships the **full feature set** — tables, sort, filter, conditional formatting, data validation, drawing/images, hyperlinks, notes, thread comments, find/replace — matching the real app. Previously the embed ran `lazyPlugins={false}` (the minimal editor) to stay a single file. But the embed's tsup build is `splitting:false` + `noExternal:/.*/`, so the lazy loader's dynamic `import()`s are **inlined** into the one `embed-runtime.js` rather than emitted as chunks — the single-file deploy is preserved. Enabling lazy plugins means the embed eager-loads any feature whose data is already in the opened file (so tables/CF are never silently dropped) and idle-loads the rest, so the toolbar/menu feature actions (Insert ▸ Table, Data ▸ Filter, …) resolve.
- dfc8e6b: Embed now honors the host's **light/dark theme**. Previously the iframe always rendered light (it never set `appearance`), so it didn't match a dark host. The runtime now reads `?theme=light|dark|system` from the embed URL, resolves `system` against the iframe's `prefers-color-scheme` (and follows live OS changes), and passes `appearance` to `<CasualSheets>` so Univer's canvas/headers/gridlines + the SDK chrome all theme together. Hosts can also push live changes over `casual.command.set.theme`.

  Also fixes a protocol bug: `EmbedHostTransport` posted `casual.command.set.{theme,readonly,locale}` (dotted) but the runtime listened for `setTheme`/`setReadOnly`/`setLocale` (camelCase), so those three host→editor commands were silently dropped. Aligned both sides to the dotted form.

## 0.13.0

### Minor Changes

- 2846392: Embed runtime `viewMode="editor"` now renders the **full SDK chrome** — the menu bar (Edit/Insert/Format/Data/View), the rich formatting toolbar (font, B/I/U, alignment, borders, number formats, colors), the formula bar, sheet tabs and status bar — so a host embeds a _complete_ spreadsheet editor and only frames/brands it, rather than hand-rolling its own toolbar.

  Previously the embed used Univer's built-in `ui` toggles, which could only show the formula bar (turning on Univer's ribbon/sheet-tabs threw `[redi]: Cannot find … registered by any injector` in the single-file bundle). It now mounts `<CasualSheets chrome="full">` — the SDK's own React chrome over the facade, which has no such service dependency and bundles cleanly into the iframe. `viewMode="preview"` stays `chrome="none"` (bare, read-only grid).

## 0.12.0

### Minor Changes

- 58ce6a0: `attachCollab` now accepts an optional `share` token (`{ token, password? }`) which is forwarded on the collab WebSocket as `?share=`/`?sp=` and suppresses the client-asserted `?role=` (the server becomes authoritative for the joiner's role). Backward-compatible: without a `share` token the connection URL is byte-identical to before. Underpins server-enforced share links (sharing-model §6.1).

### Patch Changes

- 8b35360: Two embed/chrome fixes found integrating the SDK into a host (Drive):
  - **Chrome font loader skipped Material Symbols.** `ensureChromeFonts` deduped on the bare `/css2` path, which is shared by both Google Fonts URLs, so the second family (Material Symbols Outlined) was never injected and `chrome="full"` icons rendered as raw ligature text. Now deduped per `family=` segment.
  - **`CasualSheetsIframe` ref `executeCommand` dropped `args`.** It forwarded only `{ command }` over the postMessage protocol, so iframe-host commands carrying a payload (font family/size, colour) were no-ops. Now forwards `args` too.

- 1adc983: Embed runtime `viewMode="preview"` is now genuinely READ-ONLY. Previously preview only hid the chrome (toolbar/menu) — Univer's cell editor still opened on double-click/F2, so a host's "preview" was editable.

  `applyReadOnly(univerApi, unitId, onBlock?)` now vetoes mutating commands via `beforeCommandExecuted` (throwing `CustomCommandExecutionError`, which the command service cancels cleanly). This is the load-bearing layer: the iframe's minimal plugin set does **not** enforce `WorkbookEditablePermission` (the editor still accepts edits with it flipped off), so the veto — not the permission — is what stops typing, paste, styling and structural edits. The permission flip is kept as a second layer for full `<CasualSheets>` hosts (greys out mutating menu items). The optional `onBlock(commandId)` callback lets hosts react to a blocked edit (e.g. a "read-only" toast).

  Applied in a `requestAnimationFrame` after `onReady` so it wins the race against Univer's post-mount permission init. Editor mode is unchanged. Also exposes `getEditable(univerApi, unitId)` and an `__casualEmbedApi` debug handle on the iframe window for host/e2e introspection.

## 0.11.1

### Patch Changes

- 971ad7d: Fix the in-iframe embed runtime (the `<iframe>` embed path had never been integration-tested): it mounted `<CasualSheets>` without locales, so Univer's workbench never painted (blank grid / `LocaleService: Locale not initialized`), and `embed.html` linked an `embed-runtime.css` that tsup no longer emits (now inlined), causing a 404. The embed path now boots and round-trips load → edit → save end-to-end, demonstrated by the new `examples/embed-playground`.

## 0.11.0

### Minor Changes

- 6c8a94e: Add `CasualSheetsAPI.exportXlsx()` and a `workbookDataToXlsx` converter on the `@casualoffice/sheets/xlsx` subpath — the SDK is now a two-way xlsx I/O surface (was import-only). The core converter (values/formulas, styles, merges, number formats, borders, hyperlinks, comments, data validation, tables, page setup, named ranges, VBA passthrough) was lifted out of `apps/web` and runs in its own Web Worker; ExcelJS stays out of the editor entry (lazy-loaded as a separate chunk). App-level feature models (charts/pivots/sparklines) remain a power-host concern, baked into the snapshot before serialization via the generic `ExportExtras` (`hyperlinks` / `outline` / `chartImages`).
- 3c93042: Add `CasualSheetsAPI.importXlsx(input)` — parse an `.xlsx` (`File`/`Blob`/`ArrayBuffer`/`Uint8Array`) and load it as the active workbook in one call. The ExcelJS parser is lazy-loaded from the `@casualoffice/sheets/xlsx` subpath (externalised in the build), so hosts that never import a file don't pay for it and the editor entry stays small. When a `File` is passed, its name + on-disk size are recorded on the snapshot (surfaced by the built-in Properties dialog).

## 0.10.0

### Minor Changes

- 49a3215: feat(collab): opt-in real-time co-editing via `@casualoffice/sheets/collab`

  The editor ships collab-unaware. A host enables co-editing with one call after
  `onReady`:

  ```ts
  import { attachCollab } from '@casualoffice/sheets/collab';

  const handle = attachCollab(api, { room: 'doc-42', server: 'wss://host/yjs' });
  // …later
  handle.detach();
  ```

  - `attachCollab(api, { room, server, password?, role?, token?, onSnapshot?, onStatus? })`
    spins up the Yjs doc + Hocuspocus provider + mutation bridge and returns a
    `CollabHandle` (`doc`, `provider`, `bridge`, `status()`, `detach()`).
  - The mutation bridge (`startBridge`) and replay machinery moved into the SDK —
    the non-negotiable Univer hooks (`onMutationExecutedForCollab`, `fromCollab`
    echo guard, `__splitChunk__`) travel with it.
  - `yjs` and `@hocuspocus/provider` are **peer dependencies** (optional) so the
    host provides a single Yjs copy — two copies break `Y.Doc` identity.

  Yjs/Hocuspocus is the realtime transport only; authoritative persistence stays
  host-side (WOPI / backend) via the save/exit event contract.

- 5256f3d: feat(chrome): AutoSum dropdown in the toolbar

  The built-in chrome toolbar gains an Excel-style AutoSum control (Σ): Sum /
  Average / Count numbers / Max / Min. Picking one inserts `=FN(<selection>)` one
  row below a multi-cell selection (and activates that cell), or `=FN()` into a
  single active cell. Pure facade — no Univer UI dependency — so it works in the
  embedded mount.

- 7f42243: feat(chrome): borders dropdown in the toolbar

  The built-in chrome toolbar gains a borders control (next to the colour pickers):
  a dropdown with All / Outside / Inside / Top / Bottom / Left / Right / No border.
  Each dispatches `sheet.command.set-border-position` against the active selection
  using Univer's current border style/colour. Closes a common formatting gap
  between the SDK chrome and a real spreadsheet editor.

- 29744e8: Chrome: the built-in toolbar / formula bar / status bar now flip to dark with
  `appearance="dark"`. `CasualSheets` sets the `--cs-chrome-*` CSS vars on the
  chrome wrapper from the appearance prop (hosts can still override them).
- ce87187: feat(chrome): custom Find & Replace dialog

  `<CasualSheets chrome>` now has Find & Replace, opened with Ctrl/Cmd+F (find) or
  Ctrl/Cmd+H (replace): match count + next/prev navigation (Enter / Shift+Enter),
  match-case toggle, Replace / Replace All. It's a custom, facade-driven dialog —
  search reads the active sheet's cells from `getSnapshot()`, navigation activates
  the matching cell, replace writes via `setValue` — because Univer's own
  find-replace UI doesn't render in the SDK's headless mount. Closes the last core
  chrome gap.

- 99b617f: Chrome toolbar: add font family and font size dropdowns. Apply on change via the
  Univer set-range-font-family / set-range-fontsize commands, design-token styled.
- f6b1b24: Chrome: add a minimal formula bar to `<CasualSheets chrome="minimal" | "full">`.

  Sits below the toolbar: a name box showing the active cell's A1 reference (live,
  tracks selection) and an editable input showing its formula or value. Editing
  commits through the facade — `=…` as a formula, numbers as numbers, else text.
  Self-contained (reads the active cell via `CasualSheetsAPI`, no app context, no
  autocomplete/name-box-dropdown/insert-function yet — those arrive when the rich
  `apps/web` formula bar is lifted behind `"full"`).

- 67e0d55: Chrome (`chrome="full"`): add a menu bar, color pickers, and a navigable name box.
  - **Menu bar** (Edit / Insert / Format / Data) above the toolbar — dropdown menus
    dispatching Univer commands (undo/redo, insert row/col, bold/italic/underline,
    sort asc/desc, toggle filter). No logo/title — the host frames the editor.
  - **Text & fill color pickers** in the toolbar — swatch popovers (set text color,
    fill color, or reset).
  - **Name box** in the formula bar — shows the active cell's A1 reference and jumps
    to a typed cell/range (`B5`, `A1:C3`) on Enter.

  All design-system styled, dark-mode aware, driven through `CasualSheetsAPI`.

- 7816a5d: Chrome polish: more menu items, toolbar controls, and a name-box dropdown.
  - **Menu bar**: Edit gains Cut/Copy/Paste; Insert gains Delete row/column; Format
    gains Wrap text / Clear formatting; new **View** menu (Freeze / Unfreeze panes,
    Toggle gridlines).
  - **Toolbar**: vertical align (top/middle/bottom), wrap text, clear formatting —
    with active-state reflection for vertical align + wrap.
  - **Name box**: a dropdown listing the workbook's defined names; clicking jumps to
    its range.

  All verified Univer commands, design-system styled, no new app coupling.

- 1495444: Chrome (`chrome="full"` direction): rich toolbar + design-system convergence.

  The built-in toolbar now uses Material Symbols icons (loaded idempotently via
  Google Fonts when chrome is shown) and the suite's design-system token values
  (surface-strip / text / border, light + dark). Adds strikethrough and horizontal
  alignment (left/center/right) alongside undo/redo/bold/italic/underline, grouped
  Office-style with dividers. No title/logo bar — the host frames the editor.

- 838ce1b: `<CasualSheets chrome="none" | "minimal" | "full">` — the chrome scaffold.

  First slice of the Office-chrome lift (SDK_MIGRATION_PIPELINE Phase 1 step 2).
  `chrome="none"` (default) keeps the bare grid. `"minimal"` / `"full"` wrap the
  grid in a flex column with a built-in toolbar (undo / redo / bold / italic /
  underline) that drives the editor through `CasualSheetsAPI.executeCommand` — no
  app context, no font dependency, works in any host. The rich Office shell
  (formula bar, menus, status bar) is lifted from the app behind `"full"` in later
  slices; until then `"minimal"` and `"full"` render the same toolbar.

- 3d9d0b5: Chrome toolbar now reflects the active cell: bold/italic/underline/strikethrough
  and alignment buttons light up when active, and the font family/size dropdowns
  show the current cell's font — kept in sync via a command-execution subscription.
- 35abbab: feat(chrome): sheet tabs in the built-in chrome

  `<CasualSheets chrome>` now renders a worksheet tab strip above the status bar:
  switch sheets (click), add a sheet (+), rename (double-click), and delete
  (right-click → Delete, with the last visible sheet protected). Driven entirely
  through the FUniver facade and kept live via the sheet-lifecycle events (plus a
  mutation-level fallback so collab/replay-driven changes refresh too). Closes the
  most fundamental gap between the SDK chrome and a real multi-sheet editor.

- 91ff777: Chrome: add a minimal status bar to `<CasualSheets chrome="minimal" | "full">`.

  Sits below the grid: Excel-style selection aggregates (Average / Count / Sum)
  over the numeric cells in the active multi-cell selection, live. Self-contained
  (reads the selection via `CasualSheetsAPI`). The richer status bar (configurable
  stats, min/max, zoom, sheet tabs) lifts behind `"full"`.

- f8b05b4: feat(chrome): richer status-bar stats (Numerical Count / Min / Max)

  The built-in chrome status bar now shows Excel's full selection-aggregate set —
  Average, Count, Numerical Count, Min, Max, Sum. Count is non-empty cells (any
  type); Numerical Count is numeric cells; the numeric aggregates run over numeric
  cells only, matching Excel. (A zoom control is deferred to a follow-up batch: the
  SDK's eager plugin set doesn't yet register Univer's zoom render controller.)

- a090e65: Chrome toolbar: add merge / unmerge and number-format controls (currency,
  percent, increase/decrease decimals) alongside the existing undo/redo, text
  styles, and alignment. All single-click, dispatched via Univer commands, with
  Material Symbols icons.
- 65124b4: feat(chrome): zoom control in the status bar

  The built-in chrome status bar gains a zoom control on the right: − / level / +,
  with the level click resetting to 100%. Dispatches `sheet.operation.set-zoom-ratio`
  (clamped 10–400%) — the operation path, since the higher-level zoom commands bail
  when Univer's formula-bar editor unit reports visible. Closes the zoom gap that an
  earlier batch deferred (the block was a test-timing artifact, not a real
  registration problem — Univer's zoom render controller registers in `onRendered`).

- 53b87fe: feat(collab): `attachCollab` accepts the bare `FUniver` facade, not just `CasualSheetsAPI`

  The first argument is now `CollabAttachable = CasualSheetsAPI | FUniver`. Collab
  only needs the facade, so a host that holds the raw `FUniver` (e.g. via Univer's
  own bootstrap) can attach without first wrapping it in a `CasualSheetsAPI`.
  Existing `attachCollab(api, …)` calls are unaffected.

- ea014be: Formula bar: function autocomplete. Typing a function name after `=` (e.g. `=SU`)
  shows a dropdown of matching functions; ↑/↓ to navigate, Enter/Tab to complete
  (inserts `NAME(`), Escape to dismiss. Curated common-function list.
- f0d5779: feat(sdk): `formula={{ worker }}` for off-main formula compute

  By default `<CasualSheets>` computes formulas on the main thread (fine for typical
  sheets, zero host setup). Pass a Web Worker to move compute off-thread so paste /
  sort / fill on large workbooks don't freeze the UI: the SDK registers the formula
  plugins with `notExecuteFormula` and wires `UniverRPCMainThreadPlugin` to your
  worker (dynamic-imported, so `@univerjs/rpc` stays a true optional peer — only
  loaded when a worker is passed; the host owns the worker, the SDK never bundles
  one). This is the second enabler (with `onBeforeCreateUnit`) for a power host to
  share the SDK editor core without regressing off-main compute.

- c007f64: perf(chrome): lazy-load the built-in chrome (`chrome="none"` no longer bundles it)

  `<CasualSheets>` now `lazy`-imports its chrome from the new `@casualoffice/sheets/chrome`
  subpath only when `chrome !== 'none'`. The subpath is externalised in the build, so
  the chrome stays a separate chunk the consumer's bundler code-splits — bare-grid
  hosts (the default, and any `chrome="none"` integrator) no longer carry the chrome
  JS. `dist/sheets.js` drops from ~62 KB to ~24 KB; the chrome ships as `dist/chrome.js`
  loaded on demand. The bars now appear a tick after first paint (lazy chunk load).

- 161aa91: feat(sdk): `onBeforeCreateUnit` hook to register extra Univer plugins

  `<CasualSheets onBeforeCreateUnit={(univer) => univer.registerPlugin(...)}>` fires
  after the SDK registers its built-in plugins but before the workbook unit is
  created — the only point at which a host can add register-time plugins (off-main
  formula worker via `UniverRPCMainThreadPlugin`, crosshair-highlight, zen-editor,
  …). Enables a power host to share the SDK editor core while keeping its own extra
  plugins (Phase 3). NOT semver-covered — it hands over the raw `Univer` instance.

- 3c5a990: `<CasualSheets>` save/exit events — the host-owned persistence contract (Phase 2).
  - **`onSave(snapshot)`** — fired on Ctrl/Cmd+S inside the editor (the browser save
    dialog is suppressed). The host persists the snapshot.
  - **`onExit(snapshot)`** — fired once on unmount with the final snapshot — the
    host's last chance to persist before the workbook is disposed.

  With the existing `onChange`, these complete the "the SDK emits, the host stores —
  never localStorage" model. The SDK still writes no storage of its own.

## 0.9.0

### Minor Changes

- 652068f: `CasualSheetsAPI.setTheme('light' | 'dark')` — imperative light/dark switch, the
  API equivalent of the reactive `appearance` prop. Flips Univer's
  `ThemeService.setDarkMode` (canvas colours + the `univer-dark` class Univer
  applies to the document root) via `api.setTheme(...)`, for hosts that drive the
  editor through the ref rather than re-rendering with a prop.
- f93fa6c: `<CasualSheets appearance="light" | "dark">` — reactive light/dark mode.

  Flipping it re-themes the live editor via Univer's `ThemeService.setDarkMode`
  (canvas colours, notifications, and Univer's `univer-dark` class). Distinct from
  the existing `theme` prop, which sets the Univer colour-theme object. Defaults to
  light. Note: Univer applies its dark CSS class to the document root, so dark mode
  is page-global by Univer's design.

- d3f9be6: SDK editor: working formula engine + a stable `CasualSheetsAPI` imperative ref.
  - **Formula engine now runs in embedding hosts.** The library entries
    (`index`/`sheets`/`xlsx`/`embed`/`univer`) externalise `@univerjs` so a host
    that already ships Univer no longer gets a second redi copy (which previously
    threw `[redi] loading scripts of redi more than once` and disabled the formula
    plugins). `<CasualSheets>` registers the formula engine + sheets-formula +
    numfmt and computes on the main thread.
  - **New `CasualSheetsAPI` imperative ref** handed to the host via
    `onReady(api)` — the SDK's stable integration surface:
    `getSnapshot()`, `loadSnapshot(data)`, `getSelection()`,
    `executeCommand(id, params?)`, and `api.univer` (the FUniver escape hatch,
    not covered by semver). `createCasualSheetsAPI` and the `CasualSheetsAPI` /
    `RangeRef` types are exported from `@casualoffice/sheets/sheets`.

  **Breaking:** `onReady` now receives a single `CasualSheetsAPI` argument
  instead of `(api: FUniver, univer: Univer)`. Migrate `onReady={(api) => …}`
  calls that used FUniver methods to `api.univer.<method>` (or the new
  first-class API methods where they exist, e.g. `api.executeCommand`).

  Deferred to follow-up batches: `importXlsx`/`exportXlsx` (xlsx-I/O batch),
  `setTheme` (runtime theme switch), `attachCollab` (collab adapter phase).

- 1da029e: `<CasualSheets>` now lazy-loads the feature plugins by default (`lazyPlugins`,
  default `true`): conditional formatting, data validation, hyperlinks, notes,
  tables, comments, drawings, sort, filter, and find/replace.

  Plugins whose data already lives in `initialData` (CF rules, tables, hyperlinks,
  …) load eagerly _before_ the workbook mounts, so opening a file never silently
  drops them; everything else idle-loads after first paint. This brings the SDK
  editor to feature parity with the app's grid without bloating the initial
  chunk — `@univerjs` feature packages stay external and load on demand.

  Pass `lazyPlugins={false}` for the minimal editor (render + formula + numfmt
  only); the embed-iframe runtime sets this to remain a single self-contained
  bundle.

- 2381fb4: `<CasualSheets onChange>` — a debounced stream of `IWorkbookData` snapshots.

  The "host persists it" half of the Excalidraw model: the editor stays
  storage-unaware and the host writes each snapshot wherever it likes
  (localStorage, server, …). Driven by Univer's mutation hook
  (`onMutationExecutedForCollab`), not UI events, so it captures every edit
  including programmatic ones. Debounce window is configurable via
  `onChangeDebounceMs` (default 400). Subscribed after the unit is created so
  the initial mount mutations don't emit a spurious first snapshot.

## 0.8.0

### Minor Changes

- Sheet toolbar v0.8: number formats, freeze, wrap.

  Adds to the `casual.command.execute` union:
  - `numfmt-currency`, `numfmt-percent` — single-tap apply
  - `numfmt-add-decimal`, `numfmt-subtract-decimal` — decimal stepper
  - `numfmt-custom { args.pattern }` — Excel-style pattern (e.g. `"d-mmm-yy"`, `"#,##0.00"`)
  - `wrap-toggle` — flip text wrap on the selection
  - `freeze-first-row`, `freeze-first-column`, `freeze-none` — header freezing

  All map onto canonical Univer command ids (`sheet.command.numfmt.set.currency`, `sheet.command.set-text-wrap`, `sheet.command.set-first-row-frozen`, etc.). No new format-state read-back yet — the host knows what it just dispatched, which is enough for v0.8's UX.

## 0.7.0

### Minor Changes

- Sheet toolbar v0.7: rich format commands + read-back

  Adds to the `casual.command.execute` union (host → editor):
  - `set-font-family` ({ args.family })
  - `set-font-size` ({ args.size })
  - `set-text-color` / `reset-text-color` ({ args.color })
  - `set-bg-color` / `reset-bg-color` ({ args.color })
  - `merge` / `unmerge`

  Widens `SelectionFormatStateData` (editor → host) with `fontFamily`,
  `fontSize`, `textColor`, `bgColor` read off the active cell so hosts
  can keep font / size pickers + colour swatches in sync without polling
  Univer directly.

  Fixes the v0.6 strikethrough command id (the bad `set-range-strike-through`
  which doesn't exist; the canonical id is `set-range-stroke`).

## 0.6.0

### Minor Changes

- Host-controlled toolbar wire (UX-EDITOR-1):

  New protocol envelopes
  - `casual.command.execute { command }` — host → editor. Initial union: `undo | redo | bold | italic | underline | strikethrough | align-left | align-center | align-right`. Maps to the corresponding Univer command ids inside the iframe.
  - `casual.selection.format-state { bold, italic, underline, strikethrough, align }` — editor → host. Emitted on a 200 ms poll while the workbook is mounted so hosts can mirror the active cell's format flags in their toolbar's pressed state.

  CasualSheetsIframe ref gains `executeCommand(command)`. CasualSheetsIframeProps gains `onSelectionFormatState(data)`. Drive (or any host) can now render its own toolbar above the iframe and dispatch commands without needing Univer's built-in ribbon (which the SDK can't ship because the ribbon plugins require IRPCChannelService and no worker is bundled).

  Font / size / colour / fill / merge / row+column ops are intentionally NOT in v0.6 — they need a richer command-execute payload shape we haven't locked yet.

## 0.5.7

### Patch Changes

- embed-runtime: viewMode='editor' enables Univer's formula bar + menubar (A1 ref, fx, X/✓) so the embed is visually distinct from preview mode. Toolbar + footer stay off because their workbench-mount path resolves IRPCChannelService at construction (no worker bundled). Cells remain editable via direct keyboard input.

## 0.5.6

### Patch Changes

- Drop `UniverSheetsFormulaPlugin` + `UniverSheetsFormulaUIPlugin` from
  the CasualSheets plugin chain — they resolve `IRPCChannelService`
  via Univer's DI at construction, and with no
  `UniverRPCMainThreadPlugin` registered (the SDK doesn't bundle a
  formula worker) the resolve fails with the visible console error
  "[redi]: Expect 1 dependency item(s) for id IRPCChannelService".

  Cells stay editable; formula computation is the lost capability
  (already disabled in 0.5.x via `notExecuteFormula: true`). A future
  revision can let consumers opt in to a bundled formula worker.

  Also: embed-runtime passes a UI preset to `<CasualSheets ui={...}>`
  based on the `viewMode` URL param so preview mode renders just the
  canvas. Editor mode currently uses the same preset (the toolbar
  chrome requires sheets-ui plugins not yet bundled into the embed
  runtime — tracked for 0.6.x).

## 0.5.5

### Patch Changes

- Three fixes for end-to-end iframe rendering:
  1. embed-runtime imports `../styles` (Univer CSS) so injectStyle bundles
     Univer's stylesheet into the runtime. Without this the workbench
     mounted but rendered unstyled (canvas at 0×0).
  2. Emit `parser.worker.js` directly from the embedRuntimeConfig
     (alongside embed-runtime.js) instead of relying on a post-build
     copy from mainConfig — the configs run in parallel so the copy
     races and silently fails.
  3. CasualSheets passes `notExecuteFormula: true` to the formula
     plugins so the bundle doesn't hang waiting for an
     UniverRPCMainThreadPlugin formula worker that the SDK never
     registers.

## 0.5.4

### Patch Changes

- `noExternal: ['exceljs', /^@univerjs\//]` so the parser worker
  bundles @univerjs/core (it imports LocaleType + CustomRangeType).
  0.5.3 only added exceljs to noExternal; the worker still had
  `import { ... } from "@univerjs/core"` as a bare specifier and
  closed silently at load. The "OOM" error message the embed-runtime
  emitted was misleading — it was just an unresolvable bare import
  in a module-script worker.

## 0.5.3

### Patch Changes

- Three fixes to make the iframe embed actually render:
  1. `platform: 'browser'` on the main tsup config so the parser worker
     bundles exceljs's browser fork (no Node `stream` / `buffer` / `util`
     requires that broke worker init).
  2. embed-runtime calls `transport.sendReady()` after `sendHello()`. The
     host (CasualSheetsIframe) only sends its hello inside
     `onEditorReady`; without an eager `casual.ready` from the iframe,
     the handshake deadlocked and bytes never loaded.
  3. New tsup plugin copies `dist/parser.worker.js` into `dist/embed/`
     so the `new URL('./parser.worker.js', import.meta.url)` resolution
     inside the runtime finds the worker under `{embedBasePath}/`.

## 0.5.2

### Patch Changes

- Add `platform: 'browser'` to the embed-runtime tsup config so esbuild
  picks the browser variant of dual-target deps (nanoid, etc.). 0.5.1
  bundled everything but still grabbed `import { ... } from 'crypto'`
  from the Node fork of nanoid, which the browser can't resolve. The
  runtime now lands fully clean.

## 0.5.1

### Patch Changes

- Bundle React + Univer + all deps into the embed-runtime instead of
  leaving them as external imports. The previous build expected the
  consumer to provide an importmap; consumers like Casual Drive that
  embed via `<iframe src="…/embed.html">` had no way to do that, and
  the bare `import 'react'` failed at runtime in the browser.

  The runtime now ships ~11MB self-contained (cached after first load).

## 0.5.0

### Minor Changes

- e044efd: Ship the SDK iframe-delivery architecture for sheets (Phase 2 of doc 16
  in the parent docx repo). Mirror of `@casualoffice/docs@1.1.0`.

  The existing `<CasualSheets>` direct-mount stays — no breaking change.
  Adds a new `<CasualSheetsIframe>` component that renders the editor
  inside a same-origin iframe. CSS isolation (Univer's design tokens
  stop bleeding into the host's tree), React-runtime isolation, and the
  Univer-vs-host font-cascade problems all go away when consumers
  switch from direct-mount to iframe.

  ### What the consumer-facing API looks like

  ```tsx
  import { CasualSheetsIframe } from '@casualoffice/sheets';

  <CasualSheetsIframe
    fileSource={{
      open: async (id) => ({ bytes, name, etag }), // host's bytes shim
    }}
    docId={file.id}
    viewMode="preview"             // or "editor"
    embedBasePath="/embed/sheets"   // defaults to /embed/sheets
    onSelectionChanged={…}
    onError={…}
  />;
  ```

  No iframe, no postMessage, no `EmbedTransport` wiring in the consumer.
  Bytes flow host → iframe via `casual.load.request` envelopes; inside
  the iframe the runtime parses xlsx → `IWorkbookData` via the SDK's
  own `xlsxToWorkbookData` (Phase A of #56) and mounts `<CasualSheets>`
  with the snapshot.

  ### Build artifacts

  Two new files in `dist/embed/`:
  - `embed-runtime.js` (132 KB self-contained ESM) — mounts the editor
    inside the iframe; loads xlsx bytes via the wire, converts, renders.
  - `embed.html` — the 1 KB static HTML document the iframe loads.

  Consumers copy these into their public dir at `embedBasePath` (default
  `/embed/sheets`). A Vite plugin that does the copy ships in v0.5.x;
  for v0.5.0 the contract is a two-line postinstall:

  ```sh
  mkdir -p web/public/embed/sheets
  cp node_modules/@casualoffice/sheets/dist/embed/* web/public/embed/sheets/
  ```

  ### Wire protocol additions
  - `casual.command.set.viewmode` — live preview ↔ editor toggle.
  - `casual.error` — editor → host fatal-error signal.

  Both mirror the docx repo's `13-iframe-protocol.md` extension.

  ### What's not in this minor
  - The full ref API (`flushSave`, `getSelection`, signing through iframe)
    — ships in v0.5.x once Drive proves the wire end-to-end.
  - The Vite plugin — v0.5.x.
  - Preview-mode chrome hiding inside the iframe — currently surfaced
    via `data-view-mode` attribute on the embed root + CSS gates in
    embed.html; v0.5.x wires the attribute to component-level `ui` props.
  - xlsx export from the iframe — Phase B of #56 still pending.

## 0.4.0

### Minor Changes

- Ships the xlsx **import** path as `@casualoffice/sheets/xlsx` (Phase A of [#56](https://github.com/CasualOffice/sheets/issues/56)).

  ```ts
  import { xlsxToWorkbookData } from '@casualoffice/sheets/xlsx';

  const data = await xlsxToWorkbookData(arrayBuffer);
  // → IWorkbookData ready to mount via <CasualSheets initialData={data} />
  ```

  The parser runs in a Web Worker (`parser.worker.js`, bundled as a sibling
  in `dist/`). Consumer bundlers must support the
  `new Worker(new URL(...), import.meta.url)` pattern — Vite (with
  `worker.format: 'es'`), modern webpack with worker-plugin, esbuild's
  bundler.

  ### Fidelity scope
  - Values + formulas
  - Font (family, size, bold, italic, underline, colour)
  - Fill (solid background)
  - Alignment (horizontal, vertical, wrap)
  - Number format
  - Borders (thin, per side, colour preserved)
  - Merges
  - Sheet order + names
  - Tables, comments, data validation, page setup, named ranges (resources)

  Out of scope this release: charts, drawings, pivots, sparklines,
  advanced borders (dashed/double), themes, and **export** — Phase B of
  [#56](https://github.com/CasualOffice/sheets/issues/56) handles export once
  the outline / charts / pivots / sparklines extension-point design is
  settled.

  ### What apps/web changed
  - `apps/web/src/xlsx/{import,parse-in-worker,parser.worker,parse-impl}.ts`
    and the shared utilities (`style-mapping`, `constants`, all 5
    `*-resource.ts` files, `pivot-passthrough.ts`) **moved** into
    `packages/sdk/src/xlsx/`.
  - `apps/web/src/xlsx/{export,export-impl}.ts` now imports the shared
    mappers + resource readers from `@casualoffice/sheets/xlsx`. Same
    code, new path.
  - `apps/web/src/xlsx/index.ts` re-exports `xlsxToWorkbookData` from the
    SDK so existing apps/web call-sites are unaffected.

  ### Shared internals

  The SDK's `./xlsx` entry exports the shared style mappers + resource
  readers in addition to the importer. Hosts that ship their own xlsx
  export path (Casual Sheets' apps/web is one) use them to stay in
  lockstep with this importer's shape. Consumers that only need import
  ignore them — tree-shaking strips the unused symbols.

  ### Drive unblock

  [`CasualOffice/drive`](https://github.com/CasualOffice/drive) can now replace the
  `CasualSheetWorkspace` placeholder with a real loader:

  ```tsx
  const bytes = await driveFileSource.open(file.id);
  const data = await xlsxToWorkbookData(bytes);
  <CasualSheets initialData={data} ... />;
  ```

## 0.3.0

### Minor Changes

- 73e693f: Ships `CasualSheets` — a React wrapper around Univer Sheets. Mounts a
  workbook from `initialData`, boots the eager plugin set (render +
  formula engine + UI + docs + sheets + sheets-ui + sheets-formula +
  numfmt), and surfaces the `FUniver` API to the host via `onReady`.
  Hosts (Casual Drive in particular) can now `import { CasualSheets }
from '@casualoffice/sheets/sheets'` and drop in a working
  spreadsheet view without re-implementing the boot dance.

  Lazy plugins (CF, drawings, sort, filter, hyperlinks, tables,
  comments, find/replace), the formula web worker, snapshot swap, and
  facade extensions stay app concerns — hosts layer them on top of
  `FUniver` after `onReady`.

  Also adds `./styles` (`import '@casualoffice/sheets/styles'`) as a
  side-effect entry that brings in the eager plugin CSS in one line.

  Univer 0.24.x packages move to peer dependencies (all optional, all
  declared in `peerDependenciesMeta`).

## 0.2.0

### Minor Changes

- 06a5f3a: Initial release: `@casualoffice/sheets` SDK shipping the signing pipeline
  (drawn / typed / uploaded signature surfaces, sequential / concurrent modes)
  and the iframe postMessage protocol (`EmbedTransport`, `casual.*` envelope
  types). Wire shapes are byte-identical to `@casualoffice/docs` — only
  the `app` discriminator (`'sheet'` vs `'docs'`) and signature anchor shape
  (`{ kind: 'sheet', sheet, cell }` vs `{ kind: 'doc', paraId }`) differ. The
  Univer-Sheets React wrapper (`CasualSheets` component) is planned for a
  follow-up release.
