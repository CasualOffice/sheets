# Integrating `@casualoffice/sheets`

How to embed the Casual Sheets editor in your own app. The package **is** the
editor (Excalidraw's model): you mount one React component, you own storage, and
collaboration is opt-in. There is no required backend.

This guide covers the **React component** path. For the sandboxed `<iframe>` path
(cross-origin hosts, signed embeds), see [`SDK_SIGNING_EMBED.md`](./SDK_SIGNING_EMBED.md).
For the architecture and rationale, see [`SDK_ARCHITECTURE.md`](./SDK_ARCHITECTURE.md).

> 🟢 **Published: `@casualoffice/sheets@0.10.0`** — the new Excalidraw-model editor
> SDK (full editor, `CasualSheetsAPI`, `onChange`/`onSave`/`onExit`, built-in Office
> chrome slots, `attachCollab`, lazy plugins, light/dark). It's an **early `0.x`** so
> expect additive `0.x` minors. Note the **older `@schnsrw/casual-sheets@0.8.0` is a
> different, pre-restructure line** without the API below — install
> `@casualoffice/sheets`, not `@schnsrw`.
> See [`RELEASING.md`](./RELEASING.md).

---

## Install

The library entries externalise `@univerjs/*` and `react` — your app provides a
single copy of each (bundling a second `@univerjs` copy breaks Univer's DI with
a duplicate-`redi` error). Pin **every** `@univerjs/*` package to the same
`0.25.x` version; Univer validates plugin versions against each other.

```bash
npm i @casualoffice/sheets react react-dom

# Univer peers for the eager editor (all at the SAME 0.25.x):
npm i @univerjs/core@0.25.0 @univerjs/engine-render@0.25.0 \
  @univerjs/engine-formula@0.25.0 @univerjs/ui@0.25.0 \
  @univerjs/docs@0.25.0 @univerjs/docs-ui@0.25.0 \
  @univerjs/sheets@0.25.0 @univerjs/sheets-ui@0.25.0 \
  @univerjs/sheets-formula@0.25.0 @univerjs/sheets-formula-ui@0.25.0 \
  @univerjs/sheets-numfmt@0.25.0 @univerjs/sheets-numfmt-ui@0.25.0 \
  @univerjs/themes@0.25.0
```

All `@univerjs/*` packages are declared as **optional** peers, so you install
only the ones for the features you use. The list above is the minimum the eager
editor registers (render + formula engine + UI + sheets + numfmt).

Your bundler must support the standard worker URL pattern
(`new Worker(new URL('./parser.worker.js', import.meta.url))`) for xlsx import —
Vite, modern webpack, and esbuild's bundler all do.

---

## Minimal usage

```tsx
import { CasualSheets, type CasualSheetsAPI } from '@casualoffice/sheets/sheets';
import '@casualoffice/sheets/styles';
import type { IWorkbookData } from '@univerjs/core';

export function MyEditor({ initial }: { initial: IWorkbookData }) {
  return (
    <div style={{ position: 'fixed', inset: 0 }}>
      <CasualSheets
        initialData={initial}
        onReady={(api: CasualSheetsAPI) => {
          // drive the editor imperatively from here
          console.log('ready', api.getSnapshot());
        }}
        onChange={(snapshot) => {
          // the SDK hands you the data; YOU persist it (backend, WOPI, or —
          // for a backendless host — localStorage). The SDK stores nothing.
          myBackend.autosave(snapshot);
        }}
      />
    </div>
  );
}
```

The host element must have a real size — `CasualSheets` fills its parent.
Import `@casualoffice/sheets/styles` **once** at app boot.

### A blank workbook

`initialData` is required. For an empty sheet, pass a minimal `IWorkbookData`
(one sheet, empty `cellData`), or load a snapshot you persisted earlier.

---

## Props

| Prop                  | Type                                           | Default                                           | Notes                                                                                                                  |
| --------------------- | ---------------------------------------------- | ------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------- |
| `initialData`         | `IWorkbookData`                                | —                                                 | Mounted **once**. To swap workbooks, change the React `key` to force a remount.                                        |
| `onReady`             | `(api: CasualSheetsAPI) => void`               | —                                                 | Fires after the workbook unit is created. Hands back the imperative API.                                               |
| `onChange`            | `(snapshot: IWorkbookData) => void`            | —                                                 | Debounced snapshot stream after edits settle. Driven by Univer's mutation hook, so it catches programmatic edits too.  |
| `onChangeDebounceMs`  | `number`                                       | `400`                                             | Debounce window for `onChange`.                                                                                        |
| `onSave`              | `(snapshot: IWorkbookData) => void`            | —                                                 | Fires on Ctrl/Cmd+S inside the editor (the browser dialog is suppressed). The host persists the snapshot.              |
| `onExit`              | `(snapshot: IWorkbookData) => void`            | —                                                 | Fires once on unmount with the final snapshot — the host's last chance to persist.                                     |
| `formula`             | `{ worker?: Worker \| string }`                | main-thread                                       | Off-main compute: pass a Web Worker to move formula calc off-thread (needs `@univerjs/rpc`). See **Off-main formula**. |
| `onBeforeCreateUnit`  | `(univer: Univer) => void`                     | —                                                 | Power-host escape hatch: register extra plugins before the unit mounts. NOT semver-covered. See **Power-host hooks**.  |
| `locale`              | `LocaleType`                                   | `EN_US`                                           |                                                                                                                        |
| `locales`             | `ILocales`                                     | Univer defaults                                   | String bundles. Required if you use locale-dependent UI (e.g. the formula range selector) in a non-default language.   |
| `logLevel`            | `LogLevel`                                     | `WARN`                                            |                                                                                                                        |
| `chrome`              | `'none' \| 'minimal' \| 'full'`                | `'none'`                                          | Built-in Office shell around the grid (see **Built-in chrome** below). `'none'` = bare grid (bring your own).          |
| `ui`                  | `{ header?; toolbar?; footer?; contextMenu? }` | header/toolbar/footer **off**, contextMenu **on** | Univer chrome toggles. The embedded shape hides Univer's own ribbon; build your own around the grid.                   |
| `theme`               | Univer theme                                   | `defaultTheme`                                    | Univer colour-theme **object**. Distinct from `appearance`.                                                            |
| `appearance`          | `'light' \| 'dark'`                            | `'light'`                                         | Reactive light/dark mode (`ThemeService.setDarkMode`). Univer applies its dark CSS to `<html>`, so it's page-global.   |
| `style` / `className` | —                                              | fills parent                                      | Container styling hooks.                                                                                               |
| `testId`              | `string`                                       | `casual-sheets`                                   |                                                                                                                        |

---

## Built-in chrome

By default (`chrome="none"`) the SDK renders a **bare grid** and you bring your own
toolbar/menus. Set `chrome="full"` (or `"minimal"`) to render the built-in Office
shell instead — useful for a quick full editor without building UI:

```tsx
<CasualSheets initialData={data} chrome="full" />
```

What the shell provides (all driven through the facade, no host wiring):

- **Menu bar** — Edit / Insert / Format / Data / View.
- **Toolbar** — font family + size, bold / italic / underline / strikethrough,
  text & fill colour, **borders**, horizontal & vertical align, wrap, merge /
  unmerge, number formats (currency / percent / decimals), clear formatting,
  **AutoSum** (Sum / Average / Count / Max / Min). Toggles reflect the active cell.
- **Formula bar** — name box (A1 ref + go-to + defined names) and an input with
  **function autocomplete** (`=SU` → SUM, SUMIF, …).
- **Sheet tabs** — switch (click), add (+), rename (double-click), delete
  (right-click). Last visible sheet is protected.
- **Status bar** — selection aggregates (Average / Count / Numerical Count / Min /
  Max / Sum) and a **zoom** control (− / level / +, click to reset).

Theming: every chrome part reads `--cs-chrome-*` CSS variables (set on the editor
container) and follows `appearance` for light/dark — override the variables to
match your host. `"minimal"` and `"full"` currently render the same shell;
`"full"` is where richer panels (find/replace, charts) will land.

The chrome is **lazy-loaded** (a `@casualoffice/sheets/chrome` chunk imported only
when `chrome !== 'none'`), so bare-grid consumers never bundle it — `chrome="none"`
pays nothing.

---

## Off-main formula

By default the formula engine computes on the main thread — fine for typical
sheets, zero setup. For large workbooks where paste / sort / fill would jank the
UI, move compute to a Web Worker:

```tsx
const worker = new Worker(new URL('./formula.worker.ts', import.meta.url), { type: 'module' });
<CasualSheets initialData={data} formula={{ worker }} />;
```

The SDK then registers the formula plugins with `notExecuteFormula` and wires
`UniverRPCMainThreadPlugin` to your worker. **You** own the worker (the SDK never
bundles one — brittle across bundlers) and must install `@univerjs/rpc`. The
worker script is the standard Univer formula worker (see the reference app's
`apps/web/src/univer/formula-worker.ts`).

## Power-host hooks

For hosts that want the SDK editor **core** but bring their own extra Univer
plugins (the reference `apps/web` does exactly this — it renders `chrome="none"`

- its own shell), `onBeforeCreateUnit` hands you the raw `Univer` after the SDK's
  plugins register but before the unit mounts:

```tsx
<CasualSheets
  initialData={data}
  chrome="none"
  onBeforeCreateUnit={(univer) => {
    univer.registerPlugin(UniverSheetsCrosshairHighlightPlugin);
    univer.registerPlugin(UniverSheetsZenEditorPlugin);
  }}
  onReady={(api) => {
    /* layer paste hooks / dev tools / etc. on api.univer here */
  }}
/>
```

Register-time plugins (anything that must be present when the unit inits) go in
`onBeforeCreateUnit`; everything else can layer in `onReady` via `api.univer`.
This hook is **not** semver-covered — it exposes the raw Univer instance.

---

## `CasualSheetsAPI`

The imperative ref handed to `onReady`. This is the **semver-stable** surface —
prefer it over reaching into Univer. `api.univer` is the documented escape hatch
and is explicitly **not** covered by semver.

```ts
interface CasualSheetsAPI {
  getSnapshot(): IWorkbookData | null; // current workbook
  loadSnapshot(data: IWorkbookData): void; // dispose unit + remount a new one
  importXlsx(input: ArrayBuffer | Uint8Array | Blob): Promise<IWorkbookData>; // parse + load
  exportXlsx(): Promise<Blob>; // serialize the active workbook
  getSelection(): RangeRef | null; // { unitId, sheetId, range }
  executeCommand(id: string, params?: object): Promise<boolean>;
  setTheme(appearance: 'light' | 'dark'): void; // imperative light/dark
  univer: FUniver; // escape hatch — NOT semver-covered
}
```

`importXlsx` / `exportXlsx` lazy-load the ExcelJS converters as a separate chunk
(via the `@casualoffice/sheets/xlsx` subpath), so the editor entry stays small
for hosts that never touch a file:

```ts
// open a file the user picked
await api.importXlsx(file); // File | Blob | ArrayBuffer | Uint8Array

// save the current workbook
const blob = await api.exportXlsx();
// host decides what to do with the Blob (download, upload, …)
```

### Persistence pattern — the host stores, the SDK never does

The SDK **owns no storage**. It hands you the workbook data on **change / save /
exit**; _you_ persist it wherever you want (your backend, a WOPI host, a file —
or `localStorage` if you're a backendless demo). `loadSnapshot` reads it back.

```tsx
<CasualSheets
  initialData={(await myBackend.load(id)) ?? blankWorkbook()}
  onChange={(snap) => myBackend.autosave(id, snap)} // debounced stream
  onSave={(snap) => myBackend.save(id, snap)} // explicit Ctrl+S / Save
  onExit={(snap) => myBackend.save(id, snap)} // last write before unmount
/>
```

`localStorage` is a perfectly good target **for a host that has no backend** (it's
what our Pages demo uses) — but that's _your_ choice as the host, not something the
SDK does.

In the `<iframe>` embed the _same three signals_ arrive as `postMessage`
envelopes instead of callbacks — one shape, two surfaces. Wire them with
`EmbedHostTransport` (exported from `@casualoffice/sheets/embed`):

```ts
const host = new EmbedHostTransport({ app: 'sheet', iframeWindow, embedOrigin });
host.on({
  onSaveNotify: ({ snapshot, reason }) => myBackend.save(id, snapshot), // Ctrl+S / host Save button
  onExit: ({ snapshot }) => myBackend.save(id, snapshot), // last write before the iframe unmounts
});
// onChange's iframe equivalent stays the existing selection/telemetry stream.
```

`reason` is `'shortcut'` (Ctrl/Cmd+S inside the iframe) or `'host'` (you sent
`casual.command.save` from your own toolbar). The bytes-carrying
`casual.save.request` / `onSaveRequest` pair stays available for WOPI-style
hosts that want xlsx + etag round-trips instead of a JSON snapshot.

### Reading the selection

```ts
const sel = api.getSelection();
// sel?.range → { startRow, startColumn, endRow, endColumn }
// sel?.sheetId, sel?.unitId
```

### Running commands

`executeCommand` dispatches any Univer command id (e.g. bold, undo). For the
full command set, use `api.univer` and the Univer facade.

---

## xlsx import

The pure converter lives at `@casualoffice/sheets/xlsx` and runs the parse in a
Web Worker. It works with no React and no DOM (also usable for server-side
seeding):

```ts
import { xlsxToWorkbookData } from '@casualoffice/sheets/xlsx';

const data = await xlsxToWorkbookData(await file.arrayBuffer());
api.loadSnapshot(data); // or pass as initialData on first mount
```

Fidelity: values, formulas, fonts, fills, alignment, number formats, borders,
merges, sheet order, tables/comments/data-validation/named-ranges (as
resources). Accepts loss on charts, drawings, pivots, sparklines.

---

## Collaboration (opt-in)

The editor ships **collab-unaware** — no socket, no presence, no server until
you add one. Without it you have a fully functional single-user editor.

Turn on real-time co-editing with one call after `onReady`, against a Hocuspocus

- Yjs server (the bundled `apps/server`, or your own):

```ts
import { attachCollab } from '@casualoffice/sheets/collab';

let collab;
<CasualSheets
  initialData={data}
  onReady={(api) => {
    collab = attachCollab(api, {
      room: 'doc-42',
      server: 'wss://your-host/yjs',
      // password, role: 'view' | 'write', token, onSnapshot, onStatus all optional
      onStatus: (s) => console.log('collab status:', s), // 'connecting' | 'live' | 'offline'
    });
  }}
/>;
// leaving the room (always before unmount):
collab?.detach();
```

`attachCollab` returns a `CollabHandle` — `{ doc, provider, bridge, status(), detach() }`.
The `provider.awareness` is your hook for presence (cursors, avatars); `doc` is the
raw Yjs document for any extra shared state. Build that UI on top — the SDK ships
the transport + mutation bridge, not the presence chrome.

Under the hood the bridge uses the only correct Univer hook,
`ICommandService.onMutationExecutedForCollab`, applies remote mutations with
`fromCollab` (echo-loop prevention), and guards `__splitChunk__` (see
[`CO-EDITING.md`](./CO-EDITING.md)).

> **Peer deps:** `attachCollab` needs `yjs` and `@hocuspocus/provider` from the
> host (declared `optional` peers) so there's a **single** Yjs copy in the graph —
> two copies break `Y.Doc` identity and awareness.

In collaborative mode the realtime transport carries live edits, but the
**authoritative document is persisted through your host integration (WOPI or
similar)** — the same "the host stores, the SDK doesn't" rule, not a browser
store.

---

## Versioning

`@casualoffice/sheets` (npm SDK) and the Docker app (`casualoffice/sheets`)
release on **independent** lines — see [`RELEASING.md`](./RELEASING.md). The SDK
follows Changesets; the props + `CasualSheetsAPI` are the contract, decoupled
from Univer's internal version churn. When you upgrade Univer, bump **all**
`@univerjs/*` together.

---

## Roadmap

`importXlsx` / `exportXlsx` are now on `CasualSheetsAPI` (see above); the SDK is
a two-way xlsx surface. Remaining, optional follow-ups (designed in
[`SDK_ARCHITECTURE.md`](./SDK_ARCHITECTURE.md)):

- A dedicated embed playground demonstrating the `<iframe>` path end-to-end.
- Richer chart/pivot fidelity in the core exporter (today those app-level models
  round-trip via the snapshot's resources; foreign readers see values + any
  host-supplied chart images, not live objects).
