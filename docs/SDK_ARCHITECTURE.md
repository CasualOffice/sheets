# SDK Architecture — the Excalidraw model

Target architecture for Casual Sheets as an **embeddable editor SDK + opt-in collab
server + thin reference site**. This is the *where we're going* doc; the current
runtime is described in [`ARCHITECTURE.md`](./ARCHITECTURE.md), and the staged path
between them is [`SDK_MIGRATION_PIPELINE.md`](./SDK_MIGRATION_PIPELINE.md).

> **Primary purpose of this repo:** ship a spreadsheet editor that other engines and
> apps can attach to and integrate — not a single hosted product. The hosted site is
> a reference consumer of the SDK, not the thing we build first.

---

## Reference model: how Excalidraw is structured

We are deliberately copying Excalidraw's split, because it is the cleanest known model
for "one editor, embedded everywhere, collab optional."

| Layer | Excalidraw | What it means |
| --- | --- | --- |
| **Editor package** | `@excalidraw/excalidraw` (+ `@excalidraw/element`, `/math`, `/common`, `/utils`) | The package **is the full editor** — canvas, menus, toolbar, export utils. Not a stripped core. |
| **Integration surface** | `<Excalidraw>` props (`initialData`, `theme`, `onChange`, `viewModeEnabled`…) + imperative `excalidrawAPI` ref (`updateScene`, `getSceneElements`, `exportToBlob`…) + customization children (`MainMenu`, `WelcomeScreen`, `Footer`, `Sidebar`) + standalone utils (`exportToSvg`, `serializeAsJSON`, `restore`) | Host apps drive the editor through **props + an imperative ref + slot components**, plus pure functions for import/export that need no React. |
| **App** | `excalidraw.com` | A **thin shell** (~90% the package). Static JS served by nginx. All data in `localStorage` + `IndexedDB`. The server never sees a diagram. |
| **Collab** | `excalidraw-room` (separate repo, socket.io) + Firebase for scene storage | **Opt-in.** The room server only runs when a user starts a session; payloads are end-to-end encrypted, the key lives in the URL hash and never reaches the server. |
| **Persistence** | Browser `localStorage`/`IndexedDB` by default, with a restore/migration layer | Default is zero-backend. The server is an *addition*, never a *requirement*. |

**The load-bearing idea:** the package is the editor; the app is a thin consumer of
it; collaboration and any server are strictly opt-in additions layered *around* the
package, not baked *into* it.

---

## Where Casual Sheets stands today

```
packages/sdk  →  @casualoffice/sheets (published, v0.8.0)
   ├── ./sheets        CasualSheets (MINIMAL Univer boot — core plugins only)
   │                   CasualSheetsIframe (postMessage wrapper)
   ├── ./embed         iframe postMessage protocol
   ├── ./embed-runtime self-contained in-iframe bundle
   ├── ./signing       signature pipeline
   ├── ./xlsx          ExcelJS import/export
   └── ./styles        eager plugin CSS

apps/web      →  @sheet/web (private, the REAL editor lives here)
   ├── src/UniverSheet.tsx   full editor: all lazy plugins, paste/merge hooks,
   │                         formula Web Worker, snapshot swap, dev helpers
   ├── src/shell/            Office chrome: TitleBar, Ribbon, FormulaBar, StatusBar,
   │                         FileMenu, ShareDialog …
   ├── src/file-source/      FileSource abstraction: 'browser' | 'wopi' | 'personal'
   │                         (BrowserFileSource = IndexedDB, the localStorage mode)
   └── src/collab/           Univer ↔ Yjs bridge, presence, CollabDriver

apps/server   →  @sheet/server (private, Fastify + Hocuspocus /yjs, opt-in collab)
```

### Gap analysis vs the reference model

| # | Gap | Consequence |
| --- | --- | --- |
| **G1** | **The SDK is not the editor.** `@casualoffice/sheets` boots Univer with core plugins only. The full editor — every lazy plugin, the Office chrome, paste/merge hooks, the formula worker, snapshot swap — lives in `apps/web/src/UniverSheet.tsx` + `apps/web/src/shell/` and is **not exported.** | Integrators embedding the SDK get a bare grid, not the product. They cannot reproduce the hosted editor. |
| **G2** | **No documented, stable integration API.** `CasualSheets` takes a snapshot and hands back `FUniver` via `onReady`. There is no first-class props contract (`viewMode`, `theme`, `onChange`) and no curated imperative API (`loadSnapshot`, `getSnapshot`, `exportXlsx`, `attachCollab`). | Hosts reach into raw Univer internals; every Univer bump can break them. No semver contract. |
| **G3** | **Storage + collab are tangled into the heavy app.** `FileSource` and the Yjs bridge live in `apps/web`, not as adapters around the SDK. | Collab/localStorage can't be consumed independently. An integrator who wants "editor + localStorage, no server" has to lift app code. |
| **G4** | **No thin reference site.** `apps/web` is ~60% bespoke shell; `services/site` is an Astro marketing site. The design system `@schnsrw/design-system` exists but **sheet does not consume it**. | There is no excalidraw.com-equivalent: a thin host that is mostly SDK + localStorage and demonstrates the integration path. |

---

## Target architecture

```
packages/
├── sdk            @casualoffice/sheets       ← THE EDITOR (full)
│    ├── ./           <CasualSheets> + CasualSheetsAPI imperative ref
│    ├── ./chrome      Office shell as slot components (Ribbon/FormulaBar/StatusBar/
│    │                 TitleBar/FileMenu) — opt-in, themeable, replaceable
│    ├── ./xlsx        pure import/export utils (no React)
│    ├── ./embed       iframe postMessage protocol + CasualSheetsIframe
│    ├── ./signing     signature pipeline
│    └── ./styles      CSS side-effect entry
│
├── storage        @casualoffice/sheets-storage  ← OPT-IN persistence adapters
│    ├── browser     BrowserFileSource (localStorage/IndexedDB) — the DEFAULT
│    ├── wopi        WopiFileSource
│    └── personal    PersonalFileSource (talks to apps/server)
│
└── collab         @casualoffice/sheets-collab    ← OPT-IN real-time
     ├── bridge      Univer ↔ Yjs translation + echo-loop guard
     ├── presence    cursor / selection / live-edit awareness
     └── driver      attachCollab(api, { room, server }) — connects to apps/server

apps/
├── web            @sheet/web   ← THIN reference host (excalidraw.com-equivalent)
│    └── consumes sdk + sdk/chrome + storage(browser by default) + collab(opt-in).
│        localStorage/IDB is zero-config; collab/WOPI/personal are layered adapters.
│
└── server         @sheet/server  ← OPT-IN collab + storage backend
     └── unchanged in shape: Fastify + Hocuspocus /yjs, only runs when a room exists.
```

> Package split is the *destination*. Phase 1 ships the full editor inside the existing
> single `packages/sdk` via new subpath exports; the `storage`/`collab` packages may
> start as additional subpaths (`@casualoffice/sheets/storage`, `/collab`) and graduate
> to their own packages only if a consumer needs them independently. Decide at the
> Phase 2 milestone — see the pipeline doc.

### Integration surface (the contract we owe integrators)

Modeled directly on Excalidraw's props + imperative ref + slots.

```tsx
import { CasualSheets } from '@casualoffice/sheets'
import '@casualoffice/sheets/styles'

<CasualSheets
  initialData={snapshot}          // IWorkbookData | xlsx bytes | undefined (blank)
  theme="light"                   // 'light' | 'dark'
  viewMode={false}                // read-only viewer
  chrome="full"                   // 'full' | 'minimal' | 'none'  (Office shell level)
  onChange={(snapshot) => …}      // debounced snapshot stream (host persists it)
  onReady={(api) => …}            // hands back the imperative API below
  plugins={['charts','pivot']}    // opt-in heavy plugins; default = lazy-on-demand
/>
```

```ts
interface CasualSheetsAPI {
  getSnapshot(): IWorkbookData          // current workbook
  loadSnapshot(data: IWorkbookData): void
  importXlsx(bytes: ArrayBuffer): Promise<void>
  exportXlsx(): Promise<Blob>           // pure-util path, worker-backed
  getSelection(): RangeRef
  executeCommand(id: string, params?: unknown): Promise<boolean>
  setTheme(theme: 'light' | 'dark'): void
  attachCollab(opts: { room: string; server: string; password?: string }): Detach
  univer: FUniver                       // escape hatch — explicitly "unstable"
}
```

Rules:
- **Props + `CasualSheetsAPI` are the semver surface.** `api.univer` is the documented
  escape hatch and is explicitly *not* covered by semver.
- **No backend required to embed.** `initialData` + `onChange` is enough for a host to
  run the editor against its own storage. `attachCollab` is the only thing that needs
  `apps/server`.
- **Pure import/export** (`@casualoffice/sheets/xlsx`) works with no React and no DOM —
  for server-side seeding and Node consumers (see headless caveats in `CLAUDE.md`).

### Storage: localStorage is the default, server is optional

- `BrowserFileSource` (IndexedDB, with File System Access where available) is the
  **zero-config default** — the editor persists locally with no server, exactly like
  excalidraw.com.
- `wopi` and `personal` are adapters selected at runtime by the host (`select.ts`),
  never branched on inside the editor. The editor only ever sees the `FileSource`
  interface.

### Collab: opt-in, attaches around the editor

- The editor ships **collab-unaware.** A host calls `api.attachCollab({ room, server })`
  to wire the Yjs bridge; without that call there is no socket, no presence, no server.
- Mirrors `excalidraw-room`: `apps/server` only matters once a room is created.
- The hook contract is unchanged and non-negotiable (see `CLAUDE.md` hard rules):
  subscribe to `ICommandService.onMutationExecutedForCollab`, apply remote mutations
  with `IExecutionOptions.fromCollab`, respect `params.__splitChunk__`.

> **Note (parity gap, deferred):** Excalidraw's room traffic is end-to-end encrypted
> with the key in the URL hash. Casual Sheets collab is **not** E2E-encrypted today
> because the WOPI/personal modes need server-side snapshot access. E2E is out of scope
> until a browser-only collab mode is requested; tracked as a future item, not in this
> pipeline.

---

## Design decisions

| Decision | Rationale |
| --- | --- |
| Package **is** the full editor (G1) | Excalidraw's core insight: integrators must get the product, not a kit. The hosted site then proves the embed path by *being* a consumer. |
| Props + imperative ref + slots (G2) | The exact surface Excalidraw integrators already understand; gives us a semver contract independent of Univer's internal churn. |
| Storage/collab as opt-in adapters (G3) | Lets "editor + localStorage, no server" be the default story, and keeps `apps/server` a true addition. |
| Slim `apps/web` onto the SDK (G4) | One host to maintain; it doubles as the live integration example. Adopts `@schnsrw/design-system` so all suite editors share one look. |
| Univer 0.25 first (Phase 0) | The SDK extraction should happen on the version we ship, not a moving base. The fork has no 0.25 yet — see pipeline Phase 0. |
| Keep `apps/server` shape | Hocuspocus + Fastify already match the opt-in room-server model; no redesign needed. |
