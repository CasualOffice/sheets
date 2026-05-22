# Co-editing fix tracker

Running ledger of the co-editing / cursor / sheet / border issues identified
during the May 2026 audit. One row per issue. Update **Status** and append
to **Notes** whenever a fix lands so we don't re-litigate the same bug.

Categories: `critical` blocks correctness, `high` causes silent divergence
or data loss, `medium` is visible UX rot, `low` is polish or perf.

| # | Issue | Cat | Status | Files | Notes |
|---|---|---|---|---|---|
| 1 | Password rejection silently retries (close-code 1006 vs 401) | critical | **fixed** | apps/server/src/yjs.ts, apps/web/src/collab/CollabDriver.tsx | Server now completes the WS upgrade then closes with code 4401 so the client sees a real CloseEvent. Client routes 4401 / 1008 / `authenticationFailed` to the password prompt with an error and tears down the failing provider. Regression test: `tests/e2e/coedit-regression.spec.ts`. |
| 2 | Snapshot replay race: mutations after a compaction land on old unit | critical | **fixed** | apps/web/src/collab/bridge.ts, apps/web/src/collab/CollabDriver.tsx | `onSnapshotReceived` may return a promise; the bridge `await`s it before continuing replay. `CollabDriver` awaits `requestAnimationFrame` after `replaceWorkbook` so the new unit is wired before the next op runs. |
| 3 | Op-log replay is event-sourcing, not a state CRDT | high | wontfix-v1 | apps/web/src/collab/bridge.ts | Architectural; documented in CO-EDITING.md. Mitigated by fixes 2/5/6/20. |
| 4 | No OT for structural mutations (insert-row / col races cell edits) | high | open | apps/web/src/collab/bridge.ts | Hard problem. Tracked for post-v1. |
| 5 | Silent failures: remote `executeCommand` throws are swallowed | high | **fixed** | apps/web/src/collab/bridge.ts | `.catch` added on replay; warnings now surface in console. |
| 6 | `rewriteUnitId` only patches top-level `params.unitId` | medium | **fixed** | apps/web/src/collab/bridge.ts, apps/web/src/collab/bridge-helpers.ts | New `deepRewriteUnitId` walks plain objects + arrays recursively and rewrites every `unitId` string. Pure helper extracted to `bridge-helpers.ts`; 8 unit tests pin the behaviour. |
| 7 | Seed-to-bridge gap: edits during room-create upload are lost | medium | **fixed** | apps/web/src/shell/CreateRoomDialog.tsx | Re-upload xlsx + snapshot in `openRoom()` BEFORE the navigation. Captures every edit between submit-time and Open-the-room click. Button disables during the upload with "Opening…" label. A joiner who already loaded the submit-time seed before the re-upload arrives can still be out of sync — that race needs a future "resync-on-connect" pass. |
| 8 | `params.__splitChunk__` unhandled — large pastes corrupt on peers | medium | **n/a (Univer 0.22.x)** | apps/web/src/collab/bridge.ts | `__splitChunk__` doesn't appear anywhere in our pinned Univer 0.22.x source — the CLAUDE.md hard rule was inherited from earlier Univer docs. Large pastes in 0.22.x fire one large `sheet.mutation.set-range-values` through the normal path. If a future Univer upgrade reintroduces chunked mutations, re-open. |
| 9 | 5-min room GC destroys all co-edit state | medium | **fixed** | apps/server/src/rooms.ts, apps/server/src/index.ts, docker-compose.yml | Bump default TTL from 5 → 60 minutes. Restrict eviction to "throwaway" rooms: no password AND no uploaded seed/snapshot. Password-protected rooms or rooms with content stay forever. Wire `onEvict` callback so evicted rooms also drop their persisted Y.Doc bytes from Redis (no orphan blobs). |
| 10 | Cursors render on the wrong sheet (no filter by active sheet) | medium | verify | apps/web/src/collab/PresenceLayer.tsx | Code does filter `peer.selection.s === activeSheetId`; needs e2e test to confirm. |
| 11 | Frozen panes + zoom break cursor positioning | medium | **fixed** | apps/web/src/collab/PresenceLayer.tsx, apps/web/src/charts/ChartLayer.tsx | Read freeze config via `worksheet.getFreeze()` each frame; skip the Y/X-scroll subtract for cells with row<startRow / col<startColumn. Read zoom via `worksheet.getZoomRatio()` (internal accessor — no facade getter in 0.22.x) and multiply `(content - scroll) * zoom` before adding the canvas offset + header gutter. Cursors + chart anchors now stay correct at any zoom level and through frozen quadrants. |
| 12 | liveEdit ghost jumps to next cell on Tab/Enter | low | **fixed** | apps/web/src/collab/usePresenceWire.ts | Drop the EditStarted writeLive(empty). Ghost only appears once the user actually types something (first onChanging). |
| 13 | 150 ms selection poll lags fast cursor navigation | low | **fixed** | apps/web/src/collab/usePresenceWire.ts | Subscribe to `Event.SelectionChanged` for user moves (instant); slow-poll fallback at 500ms only for programmatic moves that don't fire the event. |
| 14 | View-only users see interactive grid, edits silently vanish | medium | **fixed** | apps/web/src/collab/CollabDriver.tsx, apps/web/src/collab/view-mode.ts | Flip `WorkbookEditablePermission` to false via the injector on view-role join. Editor refuses to open and edit menu items go disabled. Re-applied on every workbook swap (snapshot replay creates a fresh unit id). Disposer restores the previous value on teardown. |
| 15 | No divergence indicator when state forks | medium | **fixed** | apps/web/src/collab/CollabDriver.tsx, apps/web/src/collab/usePresenceWire.ts, apps/web/src/shell/CollabIndicator.tsx | Each peer broadcasts a hex-encoded `Y.encodeStateVector(doc)` via awareness every 5s. CollabDriver compares local SV to every visible peer's SV every 2s. Disagreement → `syncing` for first 15s grace, then `diverged`. CollabIndicator surfaces `diverged` as an amber pulsing "Out of sync" pill. Catches stuck Yjs sync; doesn't catch the harder case of "ops applied but local Univer state forked" — that path logs via the bridge's `.catch` per #5. |
| 16 | Download returns original seed, not live state | high | **fixed (transitively via #29)** | apps/web/src/shell/file-actions.ts | The download path uses client-side `wb.save()` (reads the current Univer state). The "downloads stale data" complaint was a downstream symptom of issue #29: remote mutations weren't applying on the joiner, so the local state remained = original seed, so the export of "current state" looked exactly like the seed. With #29 fixed, the local state IS the live state and the existing download works. Verified by running coedit-cell-edit.spec.ts: after issue-29 fix, joiner B5 reads `'hello-from-alice'` straight from `wb.save()`. **Server-side `/api/rooms/:id/download` that materialises live state without joining the room is a separate "download-without-joining" feature, not a bug — deferred.** |
| 17 | Compaction `wb.save()` blocks UI thread | low | **fixed** | apps/web/src/collab/bridge.ts | `wb.save()` can't move to a Web Worker (workbook is a main-thread object graph). Switched the compaction scheduler from `setInterval` to `requestIdleCallback`; `tryCompact()` only fires when the browser reports either a forced timeout OR > 5 ms of frame budget remaining. Falls back to `setInterval` in environments without rIC. |
| 18 | `JSON.stringify` chart compare is O(n) per local update | low | **fixed** | apps/web/src/collab/CollabDriver.tsx | ChartsContext preserves per-chart references for untouched items (`prev.map(c => c.id === id ? {...c, ...patch} : c)`), so `cur !== c` is the correct diff. O(1) per chart vs O(N * payload) for the stringify-per-chart approach. |
| 19a | Canvas ID selector silently breaks on Univer upgrade | medium | **fixed** | apps/web/src/univer-dom.ts, apps/web/src/collab/PresenceLayer.tsx, apps/web/src/charts/ChartLayer.tsx | Centralised in `getUniverMainCanvas` with a fallback + one-time warning. |
| 19b | Cursors offset ~40 px due to row/col header gutter | high | **fixed** | apps/web/src/univer-dom.ts, apps/web/src/collab/PresenceLayer.tsx, apps/web/src/charts/ChartLayer.tsx | `getHeaderGutter()` returns Univer 0.22.x defaults (row 46, col 20). Initially tried the dynamic `RenderUnit.with(SheetSkeletonManagerService)` lookup but it threw `QuantityCheckError: SheetScrollManagerService not registered` when the rAF tick fired before Univer finished wiring its render-unit injector — race with first mount. Hardcoded constants trade rare custom-header support for zero-risk init. |
| 19c | Stale `rects` closure causes spurious React renders during scroll | low | **fixed** | apps/web/src/collab/PresenceLayer.tsx | Added `rectsRef` mirror; the `rectsEqual` diff now compares against the live value rather than the mount-time snapshot. |
| 19d | Cursor scroll offset double-counts frozen pane height | medium | **fixed** | apps/web/src/collab/PresenceLayer.tsx | Closed together with #11 — same `getFreeze()` check. |
| 19e | Column/row resize drag desyncs cursor by up to 67 ms | low | **fixed (transitively via #25)** | apps/web/src/collab/PresenceLayer.tsx | Once recompute runs every frame instead of every 4, resize lag drops from 67ms → 16ms (one frame). Imperceptible. |
| 20 | Remote `insert-sheet` forcibly switches local active sheet on all peers | critical | **fixed** | apps/web/src/collab/bridge.ts | Bridge captures the active sheet id BEFORE replaying any `sheet.mutation.insert-sheet` from a peer and restores it after the command resolves. Univer's `_adjustActiveSheetOnInsertSheet` no longer hijacks local view. Regression test: `tests/e2e/coedit-regression.spec.ts`. |
| 21 | `switchToSheet` bypasses command bus | medium | **fixed** | apps/web/src/shell/sheet-actions.ts | Dispatches `sheet.command.set-worksheet-activate` instead of calling `setActiveSheet` directly. |
| 22 | `renameSheet` / `hideSheet` / `showSheet` bypass command bus | medium | **fixed** | apps/web/src/shell/sheet-actions.ts, apps/web/src/collab/bridge.ts | All three now dispatch the corresponding `sheet.command.*`. `set-worksheet-hidden` added to `SYNCED_MUTATIONS` so hide/show propagate to peers. |
| 23 | Default border colour `#666` indistinguishable from gridlines | medium | **fixed** | apps/web/src/shell/home-tab-actions.ts | Now `#000000`, matching Excel default thin border. |
| 24 | Borders dropdown always uses `BorderStyleTypes.THIN` — no weight option | low | **fixed** | apps/web/src/shell/home-tab-actions.ts, apps/web/src/shell/RibbonControls.tsx, apps/web/src/shell/Toolbar.tsx, apps/web/src/styles.css | `setBorders` takes a `BorderWeight` ('thin' \| 'medium' \| 'thick' → `BorderStyleTypes.THIN/MEDIUM/THICK`). New weight-picker row in the borders dropdown, sticky for the session. |
| 25 | Remote cursor sticks to viewport during scroll — only snaps to correct cell on scroll release | critical | **fixed** | apps/web/src/collab/PresenceLayer.tsx, apps/web/src/charts/ChartLayer.tsx, apps/web/src/styles.css | Two compounding bugs: (a) CSS `transition: top/left 80ms` made every per-frame coord update lerp from the old viewport position, so during a scroll burst the cursor was always chasing; (b) recompute throttled to every 4 frames with the scroll-tick bumped INSIDE recompute, giving ~50 ms blind period at start of each scroll. Fix: drop the unconditional transition; add `.presence-cursor--moving` applied for one render after a peer's anchor cell changes (smooth cell-to-cell moves preserved); recompute every animation frame instead of every 4. Same throttle removed from ChartLayer. Per-peer anchor key (`unit:sheet:sr:sc:er:ec`) tracked in a ref so the move-vs-scroll distinction is honest. |
| 26 | Owner sees empty workbook after Share → "Open the room" | critical | **fixed** | apps/web/src/collab/CollabDriver.tsx, apps/web/src/shell/CreateRoomDialog.tsx | `window.location.href = writeUrl` is a hard navigation that nukes in-memory Univer state. The `wasOwnerOfRoom` optimisation that skipped the seed reload was wrong for any flow that goes through a real navigation — the owner ended up on an empty grid. Removed the skip; always fetch the seed/snapshot from the server on /r/<id> mount. Server cache is warm because the owner just uploaded, so the round-trip is cheap. |
| 27 | "cannot create a unit with the same unit id" on join | critical | **fixed** | apps/web/src/UniverSheet.tsx | The workbook-swap effect read `currentId` before its `await eager-plugins`, then disposed-and-created after. Two back-to-back replaceWorkbook calls (seed load → bridge snapshot replay, or any double-fire) overlapped: both async swaps awaited concurrently, both then called createUnit with the same snapshot id, second one threw. Fix: serialise swaps through a `swapChainRef` promise chain; re-read `currentId` AFTER the chain wait; defensively dispose any unit holding the target id before createUnit. |
| 28 | Password prompt appears AFTER content is loaded (security + UX) | critical | **fixed** | apps/web/src/collab/CollabDriver.tsx, apps/server/src/index.ts | Two halves: client-side, the seed/snapshot download fired before the password prompt — content rendered behind the dialog and (worse) anyone with the URL could fetch the workbook ungated. Server-side, `/api/rooms/:id/seed` and `/snapshot` had NO password check. Fix: in CollabDriver, resolve the password (stash → prompt) BEFORE any seed/snapshot fetch and pass it on `x-room-password`; in the server, gate both endpoints with constant-time `passwordOk()` via a new `checkRoomPassword` helper that accepts header OR `?p=` query. Cache-Control on snapshot becomes `private` for protected rooms. |
| 29 | **Remote mutations never apply on the joiner** — typed value seen on owner, joiner shows null | critical | **fixed** | apps/web/src/collab/bridge.ts | The `replayPending()` re-entrancy guard set `replayInFlight = (async () => { try {…} finally { replayInFlight = null } })()`. For the empty-log first call there's no `await`, so the async body ran SYNCHRONOUSLY; the `finally` set `replayInFlight = null` BEFORE the outer assignment `replayInFlight = promise` completed; the assignment then OVERWROTE the null with the now-resolved promise. From that point on every subsequent call hit `if (replayInFlight) return` and skipped the loop — remote mutations sat in the Yjs log untouched. Fix: create the promise first, assign `replayInFlight = p` BEFORE invoking the async body; clear in finally with `if (replayInFlight === p)` guard. This was the actual cause of "real-time changes propagate then revert" — the live-edit ghost cleared on commit (awareness still worked) but the underlying set-range-values mutation never applied on the receiver. Diagnosed via `tests/e2e/coedit-cell-edit.spec.ts` editor-flow regression run against the docker prod build. |

## Test coverage

| Test | Covers |
|---|---|
| `apps/web/src/collab/bridge.unit.test.ts` | `deepRewriteUnitId` — 8 cases incl. nested/array/class-instance/primitive (Fix 6) |
| `tests/e2e/coedit-regression.spec.ts` | Wrong-password rejection re-opens dialog (Fix 1); remote `insert-sheet` doesn't hijack active sheet (Fix 20) |
| `tests/e2e/coedit-share.spec.ts` | Share flow + view-only role (pre-existing) |
| `tests/e2e/coedit-compaction.spec.ts` | Op-log compaction (pre-existing) |
| `tests/e2e/coedit.spec.ts` | Basic two-peer sync (pre-existing) |

## CI

`.github/workflows/ci.yml` runs four jobs in parallel on every PR:

- **lint** — `pnpm lint` (eslint, ~30 s)
- **typecheck** — `pnpm typecheck` (tsc --noEmit across workspaces, ~60 s)
- **unit** — `pnpm test:unit` (node test runner via tsx, < 30 s)
- **e2e** — `pnpm test:e2e` (Playwright chromium, ~5 min)

## How to update this file

When a fix lands:

1. Flip **Status** to `fixed` (or `verify` if it needs an e2e to confirm).
2. Append a one-line summary to **Notes** with what changed and why.
3. If the fix is bigger than a few lines, link to the commit or PR.
4. If you discover a new symptom while fixing something else, **add a new
   row** rather than expanding an existing one — keeps history readable.

Issues marked `wontfix-v1` are deferred to a future architectural pass and
should stay in the table so a future engineer doesn't rediscover them.

## Cross-peer mutation sync gaps (audit Nov 2026)

Inventory of every `sheet.mutation.*` id Univer 0.22.x emits vs our
`SYNCED_MUTATIONS` allowlist in `apps/web/src/collab/bridge.ts`.
Generated via `grep -roh 'sheet\.mutation\.[a-z0-9-]\+' node_modules/@univerjs/*/lib/cjs/index.js`.

### Synced (will propagate to peers)

Cell-level: set-range-values, set-style. Row/col structure:
insert-row, insert-col, remove-row, remove-col, move-rows, move-cols,
set-row-hidden, set-row-visible, set-col-hidden, set-col-visible,
set-worksheet-row-height, set-worksheet-row-is-auto-height,
set-worksheet-col-width, set-row-data, set-col-data,
set-worksheet-default-style. Merges: add-worksheet-merge,
remove-worksheet-merge. Sheets: insert-sheet, remove-sheet,
set-worksheet-name, set-worksheet-order, set-worksheet-hidden,
set-tab-color. Freeze: set-frozen. Move/sort: move-range,
reorder-range. Hyperlinks: add/remove/update-hyper-link. Tables:
add-table, delete-table, set-sheet-table, set-table-filter.
Filter: set-filter-criteria, set-filter-range, remove-filter.
Notes: update-note, remove-note.

### Deferred / not yet synced

| Mutation(s) | Feature | Why deferred |
|---|---|---|
| ~~CF~~ | ~~Conditional formatting~~ | **FIXED** — mutations allowlisted + lazy-plugin gate ensures the receiver mounts the CF plugin before mutation replay. E2E: `coedit-conditional-formatting.spec.ts`. |
| ~~DV~~ | ~~Data validation~~ | **FIXED** — DV uses `data-validation.mutation.*` prefix (not `sheet.mutation.*`); allowlisted + lazy-gated. E2E: `coedit-data-validation.spec.ts`. |
| ~~set-drawing-apply~~ | ~~Insert image / shape~~ | **FIXED (transitive)** — single all-purpose mutation handles add/remove/update via JSON-1 op. Allowlisted + lazy-gated for the `drawing` plugin. Large embedded image blobs ride the op-log; acceptable until/unless a side-channel resource model is added. |
| add-range-protection, set-range-protection, delete-range-protection, add-worksheet-protection, set-worksheet-protection, delete-worksheet-protection, set-worksheet-permission-points | Range/sheet protection | Security-adjacent; needs auth-model design before propagation. |
| add-range-theme, set-range-theme, remove-range-theme, register/unregister/delete-worksheet-range-theme-style, set-worksheet-range-theme-style | Range themes | Rare power-user feature. |
| set-workbook-name, set-worksheet-column-count, set-worksheet-row-count, set-worksheet-right-to-left | Workbook/sheet metadata | Rarely changed mid-session; could add to allowlist trivially. |
| set-worksheet-row-auto-height, mark-dirty-filter-change, re-calc-filter, data-validation-formula-mark-dirty | Internal recompute markers | These are emit-side-effects of other mutations; following user-action mutations is enough. |
| toggle-gridlines, set-gridlines-color | Gridlines | Low-impact visual; could add easily. |
| update-note-position, toggle-note-popup | Note position / popup state | Cosmetic; cross-peer sync would actively be wrong (where I open a note popup is local UI state). |
| copy-worksheet-end | Internal copy-worksheet finalisation | Already covered by insert-sheet + set-range-values that precede it. |

If a deferred row becomes a user-visible problem, the fix template is the
same as for hide/show (#22): add the mutation id to `SYNCED_MUTATIONS` and,
if any React-side state subscribes to a facade event that fires only on the
COMMAND path, mirror via `CommandExecuted` for the mutation id (see
`useSheets.ts` `SHEET_LIST_MUTATIONS`).
