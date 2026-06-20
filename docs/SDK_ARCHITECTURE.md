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
| **G3** | **No clean save/exit event contract; collab is tangled into the heavy app.** The Yjs bridge + persistence live in `apps/web`, not as a host-facing event surface around the SDK. | An integrator can't get the editor to simply "hand me the data on save/exit and I'll persist it." They have to lift app code instead of consuming an event contract. |
| **G4** | **No thin reference site.** `apps/web` is ~60% bespoke shell; `services/site` is an Astro marketing site. The design system `@schnsrw/design-system` exists but **sheet does not consume it**. | There is no excalidraw.com-equivalent: a thin host that is mostly SDK and demonstrates the integration path (it may use localStorage as *its* store, like excalidraw.com — but the SDK itself stores nothing). |

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
│   (no storage package — the SDK persists NOTHING; it emits save/exit
│    events and the HOST stores the bytes. WOPI/personal/S3 adapters are
│    host-side, fed by those events.)
│
└── collab         @casualoffice/sheets-collab    ← OPT-IN real-time
     ├── bridge      Univer ↔ Yjs translation + echo-loop guard
     ├── presence    cursor / selection / live-edit awareness
     └── driver      attachCollab(api, { room, server }) — realtime transport;
                     persistence in collab mode is WOPI/host-backed, not a store

apps/
├── web            @sheet/web   ← THIN reference host (excalidraw.com-equivalent)
│    └── consumes sdk + sdk/chrome; persists via the SDK's save/exit events.
│        The Pages demo uses localStorage as ITS store (backendless, like
│        excalidraw.com); real hosts swap in WOPI / their own backend / collab.
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
  onChange={(snapshot) => …}      // debounced stream while editing (host persists)
  onSave={(snapshot) => …}        // explicit save (Ctrl+S / Save) — host writes it
  onExit={(snapshot) => …}        // editor closing — last chance to persist
  onReady={(api) => …}            // hands back the imperative API below
  plugins={['charts','pivot']}    // opt-in heavy plugins; default = lazy-on-demand
/>
```

The same events cross the iframe boundary as **postMessage** in the
`<iframe>`/`embed-runtime` delivery mode (`EmbedTransport`) — host code listens
for `save`/`exit` messages instead of React callbacks. **Persistence is always
the host's job; the SDK never writes a store.**

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

### Storage: the SDK never persists — the host does, on save/exit

- **The SDK owns no storage. No `localStorage`, no `IndexedDB`, no default store.**
  It is storage-agnostic.
- **Persistence is a handoff to the integrating app.** The editor emits the
  workbook data on **change / save / exit**; the host decides where it goes (its
  own backend, a WOPI host, a file, etc.). The editor never decides.
- **Two event-delivery surfaces, same contract:**
  - **Hooks / callbacks** for the React `<CasualSheets>` component
    (`onChange`, `onSave`, `onExit`, `onReady`).
  - **postMessage** for the `<iframe>` embed (`embed-runtime` `EmbedTransport`) —
    save/exit events cross the frame boundary to the host.
- There is **no `FileSource` baked into the editor.** A host that wants
  WOPI/personal/S3 persistence implements it on *its* side, fed by these events.

> **`localStorage` is fine for the *demo host*, not the SDK.** Our thin demo
> (`apps/web` on GitHub Pages — backendless, the excalidraw.com-equivalent) may
> persist to `localStorage` as **its** storage choice, wired through the SDK's
> save/exit events. That's the host doing it, exactly like excalidraw.com — the
> SDK still writes nothing. Real integrators swap that for their backend / WOPI.

### Collab: opt-in realtime, WOPI-backed persistence

- The editor ships **collab-unaware.** A host opts in to wire the realtime bridge;
  without it there is no socket, no presence, no server.
- **In collaborative mode, persistence is based on WOPI** (or a similar host
  protocol) — *not* a browser store. The realtime transport (Yjs/Hocuspocus)
  carries live edits; the authoritative document is saved through the host's WOPI
  integration.
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
| Host-owned persistence via save/exit events (G3) | The SDK stores nothing; it emits the data and the host persists it (localStorage for the demo, WOPI/backend for real hosts). Collab is opt-in realtime with WOPI-backed save. Keeps `apps/server` a true addition. |
| Slim `apps/web` onto the SDK (G4) | One host to maintain; it doubles as the live integration example. Adopts `@schnsrw/design-system` so all suite editors share one look. |
| Univer 0.25 first (Phase 0) | The SDK extraction should happen on the version we ship, not a moving base. The fork has no 0.25 yet — see pipeline Phase 0. |
| Keep `apps/server` shape | Hocuspocus + Fastify already match the opt-in room-server model; no redesign needed. |
