# Integrating `@casualoffice/sheets`

How to embed the Casual Sheets editor in your own app. The package **is** the
editor (Excalidraw's model): you mount one React component, you own storage, and
collaboration is opt-in. There is no required backend.

This guide covers the **React component** path. For the sandboxed `<iframe>` path
(cross-origin hosts, signed embeds), see [`SDK_SIGNING_EMBED.md`](./SDK_SIGNING_EMBED.md).
For the architecture and rationale, see [`SDK_ARCHITECTURE.md`](./SDK_ARCHITECTURE.md).

> ⚠️ **Status: in active development, not yet published.** `@casualoffice/sheets`
> is the **new** editor SDK — the Excalidraw-model restructure (full editor,
> `CasualSheetsAPI`, `onChange`, lazy plugins, …). It is **not on npm yet**; this
> guide documents the API as it lands on `main`. The **older**
> `@schnsrw/casual-sheets@0.8.0` is published but is a **different, pre-restructure
> line** — it does **not** have the API below. Don't install `@schnsrw` expecting
> this guide to work. The first `@casualoffice/sheets` publish ships the surface
> documented here. See [`RELEASING.md`](./RELEASING.md).

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

| Prop                  | Type                                           | Default                                           | Notes                                                                                                                 |
| --------------------- | ---------------------------------------------- | ------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------- |
| `initialData`         | `IWorkbookData`                                | —                                                 | Mounted **once**. To swap workbooks, change the React `key` to force a remount.                                       |
| `onReady`             | `(api: CasualSheetsAPI) => void`               | —                                                 | Fires after the workbook unit is created. Hands back the imperative API.                                              |
| `onChange`            | `(snapshot: IWorkbookData) => void`            | —                                                 | Debounced snapshot stream after edits settle. Driven by Univer's mutation hook, so it catches programmatic edits too. |
| `onChangeDebounceMs`  | `number`                                       | `400`                                             | Debounce window for `onChange`.                                                                                       |
| `locale`              | `LocaleType`                                   | `EN_US`                                           |                                                                                                                       |
| `locales`             | `ILocales`                                     | Univer defaults                                   | String bundles. Required if you use locale-dependent UI (e.g. the formula range selector) in a non-default language.  |
| `logLevel`            | `LogLevel`                                     | `WARN`                                            |                                                                                                                       |
| `ui`                  | `{ header?; toolbar?; footer?; contextMenu? }` | header/toolbar/footer **off**, contextMenu **on** | Univer chrome toggles. The embedded shape hides Univer's own ribbon; build your own around the grid.                  |
| `theme`               | Univer theme                                   | `defaultTheme`                                    | Univer colour-theme **object**. Distinct from `appearance`.                                                           |
| `appearance`          | `'light' \| 'dark'`                            | `'light'`                                         | Reactive light/dark mode (`ThemeService.setDarkMode`). Univer applies its dark CSS to `<html>`, so it's page-global.  |
| `style` / `className` | —                                              | fills parent                                      | Container styling hooks.                                                                                              |
| `testId`              | `string`                                       | `casual-sheets`                                   |                                                                                                                       |

---

## `CasualSheetsAPI`

The imperative ref handed to `onReady`. This is the **semver-stable** surface —
prefer it over reaching into Univer. `api.univer` is the documented escape hatch
and is explicitly **not** covered by semver.

```ts
interface CasualSheetsAPI {
  getSnapshot(): IWorkbookData | null; // current workbook
  loadSnapshot(data: IWorkbookData): void; // dispose unit + remount a new one
  getSelection(): RangeRef | null; // { unitId, sheetId, range }
  executeCommand(id: string, params?: object): Promise<boolean>;
  setTheme(appearance: 'light' | 'dark'): void; // imperative light/dark
  univer: FUniver; // escape hatch — NOT semver-covered
}
```

### Persistence pattern — the host stores, the SDK never does

The SDK **owns no storage**. It hands you the workbook data on **change / save /
exit**; *you* persist it wherever you want (your backend, a WOPI host, a file —
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
what our Pages demo uses) — but that's *your* choice as the host, not something the
SDK does. In the `<iframe>` embed, the same events arrive as `postMessage` instead
of callbacks.

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
you add one. Real-time co-editing is a separate service (`apps/server`,
Hocuspocus + Yjs) that attaches around the editor; the integration hook is
`ICommandService.onMutationExecutedForCollab` (see
[`CO-EDITING.md`](./CO-EDITING.md)). Without it you have a fully functional
single-user editor.

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

## Roadmap (not yet on the API)

These are designed in [`SDK_ARCHITECTURE.md`](./SDK_ARCHITECTURE.md) and land in
follow-up releases; the type never advertises a method that throws, so they
appear here, not on `CasualSheetsAPI`:

- `importXlsx` / `exportXlsx` on the API. Today: use `@casualoffice/sheets/xlsx`
  with `loadSnapshot`; the export converter is being lifted out of the host app.
- `attachCollab({ room, server })` — wire collab from the API (today: the
  `apps/server` integration above).
- `chrome="full" | "minimal" | "none"` + slotted Office chrome — the full
  ribbon/formula-bar shell promoted into the SDK.
