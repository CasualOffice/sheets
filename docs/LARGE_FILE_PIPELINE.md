# Large-File Pipeline

Plan for handling `.xlsx` files >6MB without main-thread freezes or
double-RAM footprint. Sized to be implementable in ~1 week of focused work.

Companion to [`PLAN.md`](../PLAN.md) and
[`docs/ARCHITECTURE.md`](./ARCHITECTURE.md). Out-of-scope items inherit from
[`CLAUDE.md`](../CLAUDE.md) — no persistence, no Pro, no AI.

---

## Goal

A user opens a 6–25 MB `.xlsx` and reaches "ready to edit" without the tab
appearing to freeze. The grid surface stays interactive (60 fps scroll, sub-
100 ms click latency) after open. Joining an existing room with a 6 MB+ seed
is comparable.

**Target on a Macbook M1, Chrome, 6 MB file with ~150k populated cells:**

| Stage              | Today (measured rough) | Target  |
|--------------------|------------------------|---------|
| Parse + mount      | 3–8 s, UI blocked      | <1.5 s wall, UI never blocked >100 ms |
| First paint of grid| Same as above          | <500 ms after mount |
| Save round-trip    | 4–10 s                 | <2 s    |
| Late-joiner open   | Reparse cost (3–8 s)   | <1 s    |
| Peak JS heap       | ~80–120 MB             | <50 MB  |

Numbers are targets to measure against, not promises. Stage 7 (profiling)
exists so we can prove each change moved the needle.

---

## Diagnosis

Where the time and memory actually go, with file references.

### A. XLSX parse runs on the main thread
`apps/web/src/xlsx/import.ts:72-194`
`ExcelJS.load(buffer)` + the per-cell walk all run on the main thread. For a
6 MB file this freezes the UI for several seconds. ExcelJS itself is ~600 KB
of code and sits in the main bundle just to enable this call.

### B. Joiners re-parse the same bytes
`apps/web/src/collab/CollabDriver.tsx:101-119`
Every late joiner downloads the XLSX, parses it with ExcelJS, then calls
`replaceWorkbook`. The owner already has the parsed `IWorkbookData`. The
server already has the bytes. Each new joiner pays the full parse cost.

### C. Snapshot is deep-cloned twice on save
`apps/web/src/shell/file-actions.ts:101` (`saveAsXlsx`) and `:177`
(`collectExportExtras`) both call `wb.save()`. Univer's `save()` deep-clones
the entire snapshot — see
`vendor/univer/packages/core/src/sheets/workbook.ts:132-133`. Two calls per
save = two full clones of a 6 MB tree.

### D. Snapshot held in React state alongside Univer's own copy
`apps/web/src/workbook-context.tsx` + `apps/web/src/UniverSheet.tsx:43-99`
The whole `IWorkbookData` lives in React context **and** Univer holds its own
copy after `createUnit`. Two trees on the heap for the entire session.

### E. 30+ plugins eagerly registered
`apps/web/src/univer/plugins.ts:53-104`
Drawing, notes, thread-comments, hyperlinks, tables, CF, data validation,
sort, filter, find/replace — all loaded on boot. Each adds a render
controller and per-cell interceptors that fire even when the feature is
unused. The file already flags this in its top comment.

### F. Hyperlinks replayed sequentially through the command bus
`apps/web/src/shell/file-actions.ts:241-253`
`replayPendingHyperlinks` does `await api.executeCommand(...)` once per
link. 1000 links = 1000 sequential round-trips through the worker RPC.

### G. Whole workbook built before Univer sees anything
`apps/web/src/xlsx/import.ts:119-194`
The parser walks every sheet before returning. With multiple sheets, the
user stares at the loading state until *all* are done — even though they'll
only look at sheet 1 first.

### H. Op-log grows unbounded
`apps/web/src/collab/bridge.ts:120-164`
Every synced mutation appends to one `Y.Array`. Late joiners replay
everything since room start. No compaction; no snapshotting.

### I. Implicit formula recalc setting
`apps/web/src/univer/plugins.ts:60,77`
`UniverSheetsFormulaPlugin` defaults to `initialFormulaComputing: WHEN_EMPTY`
(`vendor/univer/packages/sheets-formula/src/config/config.ts:36`). ExcelJS
emits `cell.result` for formula cells, so most cells have cached values and
this is mostly a no-op — but every formula without a cached value still
triggers a worker RPC on load. Set it to `NO_CALCULATION` explicitly.

---

## Pipeline

Seven stages. Each is independently shippable. Do them in the order listed —
earlier stages unblock later ones.

### Stage 0 — Quick-win settings (15 min)

Three config edits, no architectural change. Land these first; they cost
nothing and remove noise from the profiling baseline.

- `apps/web/src/univer/plugins.ts:60,77` — add
  `initialFormulaComputing: CalculationMode.NO_CALCULATION` to both
  `UniverSheetsFormulaPlugin` calls. Import `CalculationMode` from
  `@univerjs/sheets-formula`.
- `apps/web/src/snapshot.ts:16` — drop `INITIAL_COLUMNS` from 128 → 26.
  Growth hook (`apps/web/src/hooks/useWorkbookGrowth.ts`) handles the rest.
- `apps/web/src/shell/file-actions.ts:177` — change `collectExportExtras` to
  accept the snapshot as an argument; pass it from `saveAsXlsx` /
  `exportCurrentWorkbookAsXlsxBlob`. Eliminates the second `wb.save()` per
  export.

### Stage 1 — Parse XLSX in a worker (D1–2)

The single largest visible win. Kills (A); enables (G); shrinks the main
bundle.

**Files to add (`apps/web/src/xlsx/`):**

- `parser.worker.ts` — imports ExcelJS, accepts `{ buffer: ArrayBuffer }` via
  `postMessage`, runs the existing `xlsxToWorkbookData` body, returns
  `{ data: ImportedWorkbook }`. Use `self.postMessage(result)` with no
  transfer for the result (it's plain JSON; structured-clone handles it).
- `parse-in-worker.ts` — `parseXlsxInWorker(buffer: ArrayBuffer):
  Promise<ImportedWorkbook>`. Spawns the worker, posts the buffer with
  `[buffer]` in the transfer list (zero-copy in), awaits the result message,
  terminates the worker.

**Wiring:**

- `apps/web/src/xlsx/import.ts` — split into two modules:
  - `import.ts` keeps the pure conversion function (`workbookFromExcelJs`).
  - `parser.worker.ts` imports `workbookFromExcelJs` and ExcelJS.
  - `parse-in-worker.ts` is the main-thread entry point.
- `apps/web/src/xlsx/index.ts` — re-export `parseXlsxInWorker` as
  `xlsxToWorkbookData` so the rest of the app doesn't change.
- `apps/web/src/shell/file-actions.ts:30` (`openSpreadsheetFile`) and
  `apps/web/src/collab/CollabDriver.tsx:107` — already call
  `xlsxToWorkbookData`; nothing to change at call sites.

**Bundle hygiene:** confirm via `pnpm --filter @sheet/web build` that ExcelJS
no longer appears in the main chunk (Vite reports chunk sizes; it should now
live in `parser.worker-*.js` only).

**Exit criterion:** open a 6 MB file. Main thread should never block for
more than ~100 ms (measure with `performance.measure`). UI must remain
responsive (cursor moves, buttons click) during parse.

### Stage 2 — Stop double-parsing on join (D3)

Kills (B). Two options; we pick A for v1.

**Option A — Server caches the parsed snapshot (chosen):**

- `apps/server/src/rooms.ts` — extend the room record with
  `snapshotGzipped?: Uint8Array`.
- `apps/server/src/index.ts` — change `POST /api/rooms/:id/seed` to accept
  **two** fields via `@fastify/multipart`: the XLSX bytes (existing) and an
  optional gzipped JSON snapshot. Owner uploads both — XLSX so
  `/seed`-based download paths still work, snapshot so joiners skip parsing.
- `apps/server/src/index.ts` — add
  `GET /api/rooms/:id/snapshot` returning the gzipped JSON with
  `content-encoding: gzip`.
- `apps/web/src/collab/CollabDriver.tsx:101-119` — try
  `/api/rooms/:id/snapshot` first. If 200, stream through
  `new DecompressionStream('gzip')` (built-in to all evergreen browsers),
  `JSON.parse` the result, hand to `replaceWorkbook`. On 404 or error, fall
  back to the current `/seed` + parse path.
- `apps/web/src/shell/CreateRoomDialog.tsx` (or wherever the owner uploads
  the seed) — after parse, gzip `JSON.stringify(snapshot)` via
  `CompressionStream('gzip')` and upload alongside the XLSX.

**Why not B (Yjs-as-authority):** Cleaner long-term but requires turning the
op-log into a structured doc and is a Stage 6 concern. Don't conflate.

**Exit criterion:** opening `/r/<roomId>` for an existing 6 MB room
completes in <1 s wall clock, no ExcelJS in flight on the joiner.

### Stage 3 — Halve the in-memory footprint (D4–5)

Kills (D). Reduces peak heap by roughly the snapshot size.

- `apps/web/src/workbook-context.tsx` — change `WorkbookCtxValue` to hold
  only `{ id, sourceFormat }`. Drop `snapshot`. Add a `revision: number`
  field so swap effects can still trigger on open.
- `apps/web/src/UniverSheet.tsx:17,43` — accept `snapshot` only via a
  ref/imperative method, not as a prop tied to React state. The component
  should hold a `useRef` for the pending snapshot and consume it inside the
  swap effect, then null it out so it's GC-eligible.
- Open flow becomes: parse → push `{id, revision++, sourceFormat}` into
  context → `UniverSheet` reads `pendingSnapshot.current` → `createSheet` →
  null the ref.
- Verify with a heap snapshot in Chrome DevTools before and after: open a 6
  MB file, take heap snapshot, search for the workbook by a cell value. We
  should see one instance, not two.

**Risk:** the swap effect at `UniverSheet.tsx:77-106` is timing-sensitive
(`disposeUnit` then `createSheet`). Keep that order. Test the open-while-
in-room case (`CollabDriver.replaceWorkbook` triggers the same path).

### Stage 4 — Defer the heavy plugins (D3)

Kills (E). High impact on boot time and per-frame interceptor cost.

**Keep eager** (needed for any sheet to render):
- `UniverRenderEnginePlugin`, `UniverFormulaEnginePlugin`,
  `UniverRPCMainThreadPlugin`, `UniverUIPlugin`, `UniverDocsPlugin`,
  `UniverDocsUIPlugin`, `UniverSheetsPlugin`, `UniverSheetsUIPlugin`,
  `UniverSheetsFormulaPlugin`, `UniverSheetsFormulaUIPlugin`,
  `UniverSheetsNumfmtPlugin`, `UniverSheetsNumfmtUIPlugin`.

**Defer** (load on first use):
- CF + CF-UI → on first CF rule mount or panel open.
- Data validation (+ UI) → same.
- Hyperlink (+ UI) → when imported workbook has `__pendingHyperlinks`, or
  on first link insert.
- Note (+ UI) → on Review tab open or first note insert.
- Thread-comment + sheets-thread-comment (+ UIs) → on Review tab.
- Drawing + sheets-drawing (+ UIs) → on Insert tab or snapshot has drawings.
- Find-replace (+ sheets-find-replace) → on Ctrl+F.
- Sort, Filter (+ UIs) → on Data tab.
- Table (+ UI) → on workbook with table resources or first table insert.

**Add `apps/web/src/univer/lazy-plugins.ts`:**

```ts
type Loader = (univer: Univer) => Promise<void>;
const loaders: Record<string, Loader> = {
  conditionalFormatting: async (u) => {
    const [base, ui] = await Promise.all([
      import('@univerjs/sheets-conditional-formatting'),
      import('@univerjs/sheets-conditional-formatting-ui'),
    ]);
    u.registerPlugin(base.UniverSheetsConditionalFormattingPlugin);
    u.registerPlugin(ui.UniverSheetsConditionalFormattingUIPlugin);
  },
  // ...one per deferred group
};
const loaded = new Set<string>();
export async function ensurePlugin(univer: Univer, name: keyof typeof loaders) {
  if (loaded.has(name)) return;
  loaded.add(name);
  await loaders[name](univer);
}
```

Trigger from:
- Ribbon tab clicks (`apps/web/src/shell/MenuBar.tsx` or wherever the tabs
  are wired) — Insert tab loads drawing; Data tab loads sort+filter; Review
  tab loads notes+thread-comment.
- Snapshot inspection at open time — if `data.__pendingHyperlinks?.length`
  or `data.resources` contains CF/DV/table entries, eagerly load those
  plugins before the unit is mounted.
- Keyboard shortcuts (Ctrl+F → find-replace).

**Risk:** Univer plugin order matters (base before UI; render/formula
before sheets). Lazy loads add plugins *after* `createUnit` — verify each
deferred plugin tolerates that. The `sheets-formula` and `sheets-ui`
plugins do, per Univer's plugin lifecycle in
`vendor/univer/packages/core/src/services/plugin/plugin-holder.ts`, but the
specific plugin should be confirmed in dev with an open workbook.

### Stage 5 — Batch import side-channels (D5)

Kills (F) and partially (G).

**Hyperlinks (F):** stop using the `__pendingHyperlinks` side channel.
Encode hyperlinks directly into `IWorkbookData` cell `.p` (rich-text
`customRanges` with `rangeType: HYPERLINK`) during parse, so they ship as
part of the initial snapshot. The schema is already known — the **export**
path reads them this way at
`apps/web/src/shell/file-actions.ts:183-221` (`extractHyperlinks`). Move
that knowledge into `apps/web/src/xlsx/import.ts` (the
`pendingHyperlinks.push(...)` block at lines 159-172). On import, write the
cell's `.p` body inline; drop `__pendingHyperlinks` and
`replayPendingHyperlinks` entirely.

**Progressive sheet mount (G):** the parser worker posts the first sheet as
soon as it's ready (`{type: 'first-sheet', data: <IWorkbookData with only
sheet[0]>}`), then a final `{type: 'done', data: <full IWorkbookData>}`.
Main thread:
- On `first-sheet` → `createUnit(UniverInstanceType.UNIVER_SHEET, partial)`.
- On `done` → diff against the partial and dispatch
  `sheet.mutation.insert-sheet` for the remaining sheets.

Simpler fallback (less win, much less code): parse the whole workbook in
the worker, then on the main thread schedule additional sheets via
`requestIdleCallback`, mounting sheet 1 first.

### Stage 6 — Op-log compaction (later)

Kills (H). Not a load-time issue, but a long-lived-room cliff.

- `apps/web/src/collab/bridge.ts:120` — periodically (every N ops, or on
  the first writer's idle ticks) snapshot the workbook to a single
  `{type: 'snapshot', dataGz: Uint8Array}` record and truncate earlier
  entries. Joiners replay from the latest snapshot record onward.
- Designated "first writer" = lowest Yjs clientID present in awareness.
- Pair with Hocuspocus persistence (`onStoreDocument` hook) writing to the
  existing Redis backend (`apps/server/src/storage.ts`) so rooms survive
  restart without re-uploading the seed.

Defer this until Stage 1–5 ship and the room-longevity story actually
matters.

### Stage 7 — Profiling harness (do this BEFORE Stage 1)

Without a baseline, we can't tell if any of the above worked.

**Add `apps/web/src/perf.ts`:**

```ts
export function timeIt<T>(label: string, fn: () => T): T {
  const start = `${label}-start`;
  const end = `${label}-end`;
  performance.mark(start);
  const out = fn();
  performance.mark(end);
  performance.measure(label, start, end);
  return out;
}
export async function timeItAsync<T>(label: string, fn: () => Promise<T>): Promise<T> {
  const start = `${label}-start`;
  const end = `${label}-end`;
  performance.mark(start);
  try {
    return await fn();
  } finally {
    performance.mark(end);
    performance.measure(label, start, end);
  }
}
```

Wrap these call sites:
- `apps/web/src/xlsx/import.ts` `xlsxToWorkbookData` → `'parse-xlsx'`
- `apps/web/src/UniverSheet.tsx:43` `createUnit` → `'mount-unit'`
- `apps/web/src/UniverSheet.tsx:77-106` swap effect → `'swap-unit'`
- `apps/web/src/shell/file-actions.ts:106` `workbookDataToXlsx` →
  `'export-xlsx'`
- `apps/web/src/shell/file-actions.ts:101` `wb.save()` →
  `'snapshot-save'`

Run Chrome DevTools → Performance with the User Timing track enabled. Take
a recording on a fixed 6 MB test file before any pipeline work; save it.
Re-record after each stage. Diff is the proof.

Also add a single 6 MB fixture under `tests/fixtures/large.xlsx` if not
already present, plus a Playwright test that opens it and asserts
`parse-xlsx` measure < target. Locks in the gains.

---

## Rollout order

| Day  | Stage                                    | Effort | Risk    |
|------|------------------------------------------|--------|---------|
| 0    | Stage 7 baseline + Stage 0 quick wins    | 1 h    | None    |
| 1    | Stage 1 parser worker                    | 4 h    | Low     |
| 1    | Stage 3 part C (kill 2nd `wb.save()`)    | 30 min | None    |
| 2    | Stage 4 plugin deferral                  | 6 h    | Medium  |
| 3    | Stage 2 server-cached snapshot           | 4 h    | Low     |
| 4    | Stage 3 part D (drop React-state copy)   | 6 h    | Medium  |
| 5    | Stage 5 hyperlinks-in-snapshot           | 3 h    | Low     |
| 5    | Stage 5 progressive sheet mount (lite)   | 2 h    | Low     |
| Later| Stage 6 op-log compaction                | 1–2 d  | Medium  |

Re-record the perf trace after each stage. Stop early if measurements show
diminishing returns — not every stage is needed for every workload.

---

## Non-goals

Per `CLAUDE.md` and `PLAN.md`. Listed so we don't drift.

- **IndexedDB / localStorage autosave.** Persistence is deferred. Stage 2
  Option A keeps it server-side; Stage 6 puts it in Hocuspocus storage.
- **Patching Univer's `Workbook.save()` to use the optimized
  `cloneWorksheetData`** (`vendor/univer/packages/core/src/sheets/clone.ts:206`).
  Tempting, but a patch in `patches/` couples us to the engine internals.
  Stage 0 + Stage 3 part C eliminate the duplicate save call, which is the
  bigger win; the residual single `deepClone` is acceptable.
- **Swap ExcelJS for SheetJS.** SheetJS Community has license caveats per
  `CLAUDE.md`. ExcelJS-in-a-worker is the right call.
- **Web-worker collab driver.** Yjs and the bridge are cheap; moving them
  off the main thread won't change the wall-clock numbers.
- **Streaming XLSX parser.** ExcelJS doesn't expose one cleanly; building
  our own is weeks of work for a workload we don't have (50 MB+ files).
  Revisit only if the 25 MB upload cap rises.

---

## Open questions

Decide before starting Stage 1.

1. **Test fixture.** Do we have a canonical 6 MB+ XLSX with realistic mix
   (formulas, styles, merges, frozen panes)? If not, generate one and commit
   to `tests/fixtures/large.xlsx`.
2. **Gzip in browser.** All current evergreen browsers ship
   `CompressionStream` / `DecompressionStream`. Confirm browser-targets
   document (Phase 0 open question in `PLAN.md`) hasn't been narrowed to
   exclude any.
3. **Server payload cap.** `apps/server/src/index.ts:18` sets multipart
   `fileSize: 25 MB`. Stage 2 adds a gzipped snapshot upload — typically
   ~20% of the XLSX size after gzip, but for very text-heavy sheets it can
   exceed it. Bump the cap to 50 MB when Stage 2 lands, or stream-write the
   snapshot to disk instead of buffering.