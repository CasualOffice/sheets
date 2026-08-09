# Casual Sheets SDK — consumer guide

`@casualoffice/sheets` embeds a full spreadsheet editor (Univer Sheets under the
hood) into your app. The host owns storage; the SDK stays storage-unaware and
hands you content snapshots to persist. This guide covers the in-process React
surface: mounting, load/save, events, the imperative handle, chrome, collab, and
AI.

> The docs SDK (`@casualoffice/docs`, `CasualEditor`) exposes the **same unified
> contract** — same `documentMode`, `on/off` events, `getContent`/`setContent`,
> `features`, `extensions`, and `collab` shapes — so a host wires both editors
> the same way. The canonical spec is doc 38 (unified SDK contract).

## Install

```bash
npm install @casualoffice/sheets
```

Import the plugin CSS once at app boot (side-effect entry):

```ts
import '@casualoffice/sheets/styles';
```

## Mount

`<CasualSheets>` mounts a single workbook from an `IWorkbookData` snapshot and
hands you the imperative API through `onReady`.

```tsx
import { CasualSheets, type CasualSheetsAPI } from '@casualoffice/sheets';
import '@casualoffice/sheets/styles';

function Editor({ initialData }) {
  return (
    <CasualSheets
      initialData={initialData}
      onReady={(api: CasualSheetsAPI) => {
        // stash the handle; drive the editor through it
      }}
      onChange={(snapshot) => persist(snapshot)}
      style={{ width: '100%', height: '100%' }}
    />
  );
}
```

`initialData` is read **once** on mount. To swap workbooks, either change the
React `key` to remount, or call `api.setContent(data)` on the handle.

### Iframe mount (cross-origin / sandboxed hosts)

For hosts that prefer process isolation, `CasualSheetsIframe` renders the editor
in an iframe and talks to it over the `embed` postMessage protocol:

```tsx
import { CasualSheetsIframe, type CasualSheetsIframeRef } from '@casualoffice/sheets';
```

## Document mode

`documentMode` is the SuperDoc-aligned interaction mode, shared with the docs
SDK. Sheets supports two values:

```ts
type DocumentMode = 'editing' | 'viewing';
```

- `'editing'` (default) — fully editable.
- `'viewing'` — read-only (command-veto + `WorkbookEditablePermission` path).

It is reactive — flipping the prop re-applies via `api.setDocumentMode`. The
deprecated `readOnly` boolean maps to `'viewing'` only when `documentMode` is
unset; `documentMode` always wins.

```tsx
<CasualSheets initialData={data} documentMode="viewing" />
```

## Load & save

The host owns persistence. Read and replace content via the handle:

| Method             | Purpose                                                      |
| ------------------ | ------------------------------------------------------------ |
| `getContent()`     | Current workbook as `IWorkbookData` (or `null` pre-mount).   |
| `setContent(data)` | Replace the workbook with a new snapshot; clears dirty flag. |
| `import(input)`    | Parse an `.xlsx` `ArrayBuffer`/`Uint8Array`/`Blob` and load. |
| `export()`         | Serialize the workbook to an `.xlsx` `Blob`.                 |

```ts
const snapshot = api.getContent();
api.setContent(nextSnapshot);

const blob = await api.export(); // .xlsx Blob
await api.import(fileFromInputElement); // File / Blob / ArrayBuffer
```

`import`/`export` are the canonical cross-editor aliases of the format-specific
`importXlsx`/`exportXlsx` (both exist; the ExcelJS converter is lazy-loaded as a
separate chunk). `getSnapshot`/`loadSnapshot` are deprecated aliases of
`getContent`/`setContent`.

The SDK also fires an explicit `onSave` when the user presses `Ctrl/Cmd+S`
inside the grid (the browser dialog is suppressed), and `onExit` once with the
final snapshot on unmount — the host's last chance to persist.

## Events

Every canonical event is available **two ways**: as an `on*` prop on
`<CasualSheets>` **and** via `api.on(name, handler)` / `api.off(name, handler)`
on the handle. Same event, same payload. `api.on` returns an unsubscribe
function; `'ready'` is sticky (a late subscriber fires immediately).

| Event / prop                            | Payload                         |
| --------------------------------------- | ------------------------------- |
| `ready` / `onReady`                     | `(api: CasualSheetsAPI)`        |
| `change` / `onChange`                   | `(snapshot: IWorkbookData)`     |
| `selectionChange` / `onSelectionChange` | `(selection: RangeRef \| null)` |
| `save` / `onSave`                       | `(snapshot: IWorkbookData)`     |
| `error` / `onError`                     | `(error: Error)`                |
| `dirtyChange` / `onDirtyChange`         | `(dirty: boolean)`              |

`onChange` is debounced (default 400 ms; tune with `onChangeDebounceMs`) and
driven by Univer's mutation hook, so it captures programmatic edits too.
`dirtyChange` flips `true` on the first edit since the last load/save and
`false` on save / `setContent` / `import`.

```ts
const off = api.on('selectionChange', (sel) => updateStatusBar(sel));
// later
off();
```

## Imperative API (`CasualSheetsAPI`)

The handle from `onReady` is the stable, semver-covered integration surface.

| Method                                | Purpose                                                       |
| ------------------------------------- | ------------------------------------------------------------- |
| `getContent()`                        | Current `IWorkbookData` snapshot, or `null`.                  |
| `setContent(data)`                    | Replace the workbook; clears the dirty flag.                  |
| `import(input)` / `importXlsx(input)` | Load an `.xlsx` (`ArrayBuffer` / `Uint8Array` / `Blob`).      |
| `export()` / `exportXlsx()`           | Serialize the workbook to an `.xlsx` `Blob`.                  |
| `getSelection()`                      | Active selection as a `RangeRef`, or `null`.                  |
| `focus()`                             | Move keyboard focus to the active workbook.                   |
| `on(name, handler)`                   | Subscribe to an event; returns an unsubscribe fn.             |
| `off(name, handler)`                  | Remove an event handler.                                      |
| `executeCommand(id, params?)`         | Dispatch a Univer command; resolves to its boolean result.    |
| `executeCommands(steps)`              | Replay a sequence of recorded command/mutation steps.         |
| `onMutation(handler)`                 | Observe the replayable mutation stream; returns a disposer.   |
| `onLocalMutation(handler)`            | Observe locally-authored mutations for persistence.           |
| `setTheme(appearance)`                | Imperative light/dark switch (`'light'` \| `'dark'`).         |
| `setDocumentMode(mode)`               | Switch between `'editing'` and `'viewing'`.                   |
| `getDocumentMode()`                   | The current `DocumentMode`.                                   |
| `univer`                              | Raw FUniver facade — escape hatch, **not** covered by semver. |

`getSnapshot` / `loadSnapshot` remain as deprecated aliases of
`getContent` / `setContent`.

```ts
await api.executeCommand('sheet.command.set-range-values', { value: 42 });
const range = api.getSelection(); // { unitId, sheetId, range }
```

## Chrome & features

By default the SDK renders a **bare grid** (`chrome="none"`) so the host brings
its own shell. Pass `chrome="minimal"` or `chrome="full"` to get the built-in
Office shell (menu bar, formatting toolbar, formula bar, sheet tab strip, status
bar).

```tsx
<CasualSheets initialData={data} chrome="full" features={{ merge: false }} />
```

### `features` map

`features?: Record<string, boolean>` toggles individual chrome controls. `false`
hides the control **and**, for command-backed capabilities, vetoes the matching
Univer command at the shared command bus. The veto still applies with
`chrome="none"`, so a shortcut, context menu, facade call, `api.executeCommand`,
`api.executeCommands`, or the raw `api.univer` escape hatch cannot bypass it.
Omitted, `true`, and unknown keys stay enabled.

`api.univer.executeCommand` and `syncExecuteCommand` deliberately discard
caller-supplied collaboration, snapshot, formula-result, and transport-only
provenance (`fromCollab`, `fromChangeset`, `fromFormula`,
`applyFormulaCalculationResult`, `onlyLocal`, and `syncOnly`). Those flags are
engine attribution, not a permission or persistence override. The SDK's
collaboration bridge and formula engine use private command-service seams so
authoritative remote changes and recalculation results still converge.
Univer's private `_injector` field is unsupported SDK-internal machinery;
hosts must not use it as an integration or policy-bypass surface.

`onMutation` retains its historical all-mutation audit/replay behavior,
including remote and engine-local mutations. Use `onLocalMutation` for a host
save queue: it excludes collaboration/changeset replay and `onlyLocal` engine
writes such as formula-result caches. Public command dispatch cannot forge
those exclusions because it strips the corresponding execution markers.
Exceptions from either observer are emitted through the SDK `error` event and
never change the result of an already-applied engine mutation.

Known compound commands are preflighted before Univer starts their nested
mutations. Paste, clear, format painter, and auto-fill are rejected at their
root when a disabled downstream capability could otherwise veto after earlier
mutations and leave an un-undoable partial edit. Explicit value/formula-only
and column-width paste modes remain available when their own capabilities are
enabled. Structural row/column/range operations and sheet copy/removal are
likewise preflighted against plugin-maintained tables, conditional formatting,
filters, data validation, and merges before the sheet shape changes.
Optional-paste replacement, refill, undo, and redo are also blocked while a
capability they may replay is disabled because their prior payload is held in
engine history rather than available to the policy for safe inspection.

Every nested command is still evaluated independently. Univer's execution
stack is global rather than async-context-local, so lending a root decision to
apparent descendants would let an unrelated concurrent UI/API command inherit
the wrong authority. A host plugin that implements another multi-mutation or
async compound action must therefore expose and preflight its own root command;
it must not rely on a late child veto to roll back earlier mutations.

Command-backed feature ids are `history`, `clipboard`, `format-painter`, `font`,
`font-style`, `color`, `borders`, `alignment`, `merge`, `number`, `clear-format`,
`tables`, `conditionalFormatting`, `filter`, and `dataValidation`. UI-only or
non-dedicated-command ids — `file`, `help`, `branding`, `autosum`, `format-cells`,
`insert-chart`, `pivot-table`, `charts`, `pivots`, and `sparklines` — remain
surface gates only. They do not claim engine enforcement; use `onBeforeCommand`
when a host plugin gives one of those resources its own command id.

### `onBeforeCommand` policy

Use the synchronous `onBeforeCommand` prop for host-specific policy. Return
`false` to cancel a local engine command before its handler runs:

```tsx
<CasualSheets
  initialData={data}
  onBeforeCommand={({ id }) => (id === 'sheet.command.delete-range' ? false : undefined)}
/>
```

The callback sees local commands from every entry path, including nested
commands and mutations. Authoritative collaboration, snapshot-load replays,
and internal formula-result writes bypass local policy so a restricted client
still converges and recalculates. Feature ids are per-format (a sheet's `merge`
capability differs from a doc's `trackChanges`).

Hosts implementing a narrower editor can import the SDK's pure
`isReadOnlyBlocked(commandId)` predicate and allow only their explicitly
supported edit commands. The predicate is intentionally conservative: known
navigation/selection operations and copy remain available, while commands and
mutations otherwise fail closed, including future plugin ids. This keeps the
read-only boundary in one place instead of copying Univer command-id heuristics
into each host.

### `extensions` (ChromeExtensions slots)

`extensions?: ChromeExtensions` lets a host append custom UI and override
dialogs on top of the built-in chrome (`chrome="full"`). Every field is
optional; built-ins are the defaults, host entries append or override.

```ts
interface ChromeExtensions {
  toolbar?: ToolbarExtension[]; // id, label, icon, onClick(api) | command, isVisible(api)
  menu?: MenuExtension[]; // menu target, id, label, icon?, shortcut?, onClick | dialog
  dialogs?: Partial<Record<DialogKind, DialogExtension>>; // add or OVERRIDE a dialog by kind
  panels?: PanelExtension[]; // id, title, railIcon, component
}
```

Every extension is handed the live `CasualSheetsAPI`; dialog/panel components
also receive an `onClose`. A host toolbar/menu item either runs `onClick(api)`
or dispatches a Univer `command`. `MenuTarget` is one of `file`, `edit`, `view`,
`insert`, `format`, `data`, `help`.

```tsx
<CasualSheets
  initialData={data}
  chrome="full"
  extensions={{
    toolbar: [
      {
        id: 'export-pdf',
        label: 'Export PDF',
        icon: 'picture_as_pdf',
        onClick: (api) => exportPdf(api.getContent()),
      },
    ],
  }}
/>
```

## Collab (declarative)

Pass a `collab` prop to join a real-time room; the SDK wires Yjs/Hocuspocus
itself once the editor is ready and detaches on unmount (re-attaching when
`server` / `room` / `password` / `token` / `role` change). Omit it for a
single-user editor.

```tsx
<CasualSheets
  initialData={data}
  collab={{
    server: 'wss://your-host/yjs',
    room: 'workbook-42',
    token: authToken, // optional; defaults to 'anon'
    role: 'write', // 'view' | 'write'; default 'write'
    compaction: 'off', // use backend-owned validated checkpoints
    onStatus: (s) => setStatus(s), // 'connecting' | 'live' | 'offline'
    onSnapshot: (wb) => {}, // peer compaction snapshot arrived
  }}
/>
```

Options match `AttachCollabOptions` (`server`, `room`, `password`, `token`,
`role`, `share`, `onStatus`, `onSnapshot`, `getContent`, `compaction`). Pass
`compaction: 'off'` when the host/server owns checkpoints; this is also the safe
default. `'manual'` keeps only the bridge's health-gated `forceCompact()`, while
`'auto'` explicitly enables the idle timer. A full
`CasualSheetsAPI` automatically supplies preservation-aware snapshot I/O. A bare
FUniver host must supply `getContent` plus an awaited `onSnapshot` or concurrent
reorder recovery and compaction fail closed. Yjs/Hocuspocus is the realtime
transport only — the authoritative document is still saved by the host via the
save/exit contract; collab does not turn the SDK into a store.

For advanced hosts that drive the room lifecycle themselves (presence UI,
preflight, reconnect banners), the imperative `attachCollab(api, opts)` export
from `@casualoffice/sheets/collab` stays available — don't combine both on one
editor.

## AI (declarative)

Pass an `ai` prop to mount a task-pane surface beside the grid. The SDK owns the
prop contract, the `SheetsAiTransport` type, and the layout slot; the host
supplies the panel body via `ai.render` and drives its tool loop against
`ai.transport`.

```tsx
import { CasualSheets, createSheetsAiTransport } from '@casualoffice/sheets';

<CasualSheets
  initialData={data}
  ai={{
    enabled: true,
    transport: createSheetsAiTransport(), // desktop-native → collab → browser-direct
    render: (ctx) => <YourAiPanel {...ctx} />, // ctx: { api, transport, onAction, close }
    onAction: (action) => trackAiAction(action),
  }}
/>;
```

`enabled` is reactive (mounts/unmounts the pane once ready). `createSheetsAiTransport()`
picks the transport for the environment: `DesktopAiTransport` inside the desktop
shell, `CollabAiTransport` when a collab WS URL is available (server holds the
key), else `DirectAiTransport` (bring-your-own Anthropic key). Omit `ai` for an
editor with no AI.

## Appearance

`appearance?: 'light' | 'dark'` is reactive — flipping it re-themes the live
editor via Univer's `ThemeService`. The imperative equivalent is
`api.setTheme('light' | 'dark')`.
