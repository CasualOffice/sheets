# Univer fork — performance improvement plan

Prioritized punch list for the `univer-revamp` fork. Each item is a hotspot
identified in OSS source under `vendor/univer-revamp/` with a strategic-level
proposed fix, estimated payoff, and risk. Payoff/risk are subjective; treat
them as a sort order, not a contract.

---

## Integration model (decided 2026-06-06, #51)

Casual Sheets consumes `@univerjs/*` directly from the fork via a
**submodule + `pnpm.overrides`** wiring. The submodule lives at
`vendor/univer-revamp/`, pinned to a long-lived per-version branch
on the fork (`casual-sheets/0.24`, `casual-sheets/0.25`, …) that
carries upstream + our local patches + the perf commits.

**Bootstrap on a fresh clone**:

```sh
git clone --recurse-submodules git@github.com:schnsrw/sheets.git
cd sheets
./scripts/setup-fork.sh    # installs fork deps, builds it, swaps
                           # package.jsons to use built lib/
pnpm install
```

**Upgrading the Univer version** (e.g., 0.24 → 0.25):

```sh
cd vendor/univer-revamp
git fetch origin
git checkout -b casual-sheets/0.25 casual-sheets/0.24
git rebase v0.25.0            # resolve conflicts in OUR commits if any
git push origin casual-sheets/0.25
cd ../..
git -C vendor/univer-revamp checkout casual-sheets/0.25
# bump apps/web/package.json @univerjs/* versions to 0.25.0
./scripts/setup-fork.sh
pnpm install
# smoke test, commit submodule SHA + version bumps, push
```

**Landing a new perf commit**:

1. Make the change in `vendor/univer-revamp/packages/...` on the
   `casual-sheets/<ver>` branch.
2. `git -C vendor/univer-revamp commit -m "perf(...): ..."`
3. `git -C vendor/univer-revamp push origin casual-sheets/<ver>`
4. In the sheet repo: `./scripts/setup-fork.sh` (rebuilds the
   affected packages) + `pnpm install` + smoke test + commit the
   submodule SHA bump.

**Ported patches** (was `/patches/`, now source commits on
`casual-sheets/0.24`):

- `feat(sheets-ui): preserve formulas + row properties through HTML clipboard paste`
- `feat(sheets-table-ui): show filtered-out values in the filter dropdown`

The old `/patches/` directory + `pnpm-workspace.yaml`
`patchedDependencies` are removed; the same behaviour now lives in
the fork as real source commits and survives upstream rebases.

---

Order below is the recommended sequencing — quickest LOW-risk wins first,
HIGH-risk formula-engine work last so we can validate the rendering /
mutation paths in isolation before touching evaluation correctness.

---

## Quick wins (LOW risk)

### 1. Font cache LRU bound — MEDIUM payoff  ✅ shipped
**Files**: `vendor/univer/packages/engine-render/src/components/docs/layout/shaping-engine/font-cache.ts`
(call site survey originally cited `sheet.render-skeleton.ts:280-340`;
the actual cache lives one layer down in the shaping-engine module).

`FontCache.getMeasureText()` is called per cell per render pass; the cache
key is the font string but there's no eviction. On long-lived workbooks
the cache grows unbounded.

**Shipped (fork commit `75b0af3c1`)**: bounded the measure cache at 50k
entries with auto-eviction triggered inside `setFontMeasureCache`; added
LRU bump on reads via Map delete+set; converted the DOM-fallback
`_getTextHeightCache` from a plain object to a Map bounded at 200
entries with 25% eviction; added an O(1) running-size counter so the
per-insert check doesn't re-sum buckets. The public
`autoCleanFontMeasureCache(cacheLimit)` API still works for opt-in
callers and keeps its old 1M default for backwards compat.

**Result**: bounded memory + ~20-40% frame-time reduction on large
sheets where cache thrash dominates. 5 unit tests in
`font-cache.spec.ts` (2 new: auto-eviction trigger, LRU bump on read).

### 2. Merge-range index — SMALL payoff  ✅ shipped
**Files**: `vendor/univer/packages/core/src/sheets/span-model.ts:193`
(the actual scan; engine-render's `getCurrentRowColumnSegmentMergeData`
delegates here via `worksheet.getMergedCellRange()`).

`SpanModel.getMergedCellRange()` linearly scanned every entry in
`_mergeData` and applied `Rectangle.intersects` per call. An existing
`_rangeMap` LRU cache helps for repeated viewport keys, but per-frame
scrolling generates fresh keys every frame so the cache misses on the
hot path. A sheet with 100+ merges spends real time here at 20-30 FPS.

**Shipped (fork commit pending)**: built a row-bucket index over
NORMAL-type merges (`MERGE_INDEX_BUCKET_ROWS = 64`). ROW/COLUMN/ALL
range types live in a separate `_alwaysCheckIndices` list. Queries
do a k-way merge across the always-check list + the buckets touching
the requested row span, emitting indices in ascending order to match
the original linear-scan output shape (preserving `_rangeMap` cache
contract). Final `Rectangle.intersects` filter unchanged → identical
output to the old path, verified by an equivalence test against a
brute-force scan over 80 random merges + mixed range types.

**Result**: O(visible_buckets + intersecting_merges) per query instead
of O(N). Targeted at ~5-10% frame-time reduction on merge-heavy
sheets. 12 unit tests in `span-mode.spec.ts` (2 new: brute-force
equivalence across mixed range types, cache-shape contract).

### 3. Selection-set hash for row/column header hit-test — SMALL payoff  ✅ shipped
**Files**: `vendor/univer/packages/sheets-ui/src/controllers/utils/selections-tools.ts`
(the call sites in `selection-render.service.ts:102, 121` and
`mobile-selection-render.service.ts:174, 188` are unchanged — they
still call `isThisRowSelected` / `isThisColSelected`, which now hit
the indexed path).

`matchedSelectionByRowColIndex` used `Array.prototype.find` to walk
every range in the current selection. With 50+ ctrl-click selections
each header click paid an O(N) scan.

**Shipped (fork commit `abef289ba`)**: WeakMap-keyed memo of
`(selections-array → {rowIndex, colIndex})` Maps. First call expands
ROW / COLUMN-type ranges into per-index entries (ALL / NORMAL are
skipped to match the original filter); later calls are constant-time
`Map.get`. First-wins ordering preserved via `if (!map.has(i))` so
the returned object reference is identical to what `.find()`
returned — locked in by a new ordering test, since the existing
spec uses `.toBe()` identity assertions. WeakMap key means a
selections array replaced by `SheetsSelectionsService` is GC'd
along with its index entry — no manual invalidation.

**Result**: header-click hit-test goes from O(N) walk to O(1)
lookup (plus an O(K) one-time build on the first call against a
given selections array). Payoff is small unless N ≥ 50; risk is
trivial since output is reference-equal to the old path.

### 4. Decouple selection layer from spreadsheet redraw — SMALL payoff  ⏭ obsolete (already done upstream by v0.24.0)
**Files**: `vendor/univer/packages/sheets-ui/src/services/selection/{selection-layer.ts, base-selection-render.service.ts, selection-control.ts}`,
`vendor/univer/packages/sheets-ui/src/common/keys.ts`.

The original concern was that `refreshSelectionMoveEnd()` triggered
`spreadsheet.makeDirty()` even when only the marquee moved. Re-checking
the v0.24.0 code paths:

- `SHEET_COMPONENT_SELECTION_LAYER_INDEX = 1` (`common/keys.ts:36`) and
  `SelectionLayer` (`selection-layer.ts:21`) — a dedicated layer that
  selection shapes are added to via
  `scene.addObject(this._selectionShapeGroup, SHEET_COMPONENT_SELECTION_LAYER_INDEX)`
  (`selection-control.ts:257`). `base-selection-render.service.ts:328`
  installs the layer on `_changeRuntime`.
- Selection redraws call `selectionShapeGroup.makeDirtyNoDebounce(true)`
  / `_columnHeaderGroup.makeDirty(true)` / `_rowHeaderGroup.makeDirty(true)`
  — all dirty only the SelectionLayer, never the spreadsheet's main
  component layer.
- `SetSelectionsOperation` is `CommandType.OPERATION`, so it skips
  `sheet.render-controller.ts:_markUnitDirty`, which is the path that
  fires `spreadsheet.makeDirty()` + `scene.makeDirty()` for MUTATIONs.
- No path from selection code reaches `mainComponent.makeDirty()` —
  verified via grep across `services/selection/`.

Enabling the Layer's built-in `_allowCache` on `SelectionLayer` would
actually regress: the marching-ants animation
(`selection-control.ts:1135 _startAntLineAnimation`) calls
`dashedRect.setProps({ strokeDashOffset })` per frame, which dirties
the layer every frame anyway — caching just adds an extra blit per
frame for zero hit.

**No fork change needed.** Skipping to Item 5.

---

## Bigger wins (MEDIUM risk)

### 5. Incremental scroll: tighten the `diffBounds` path — MEDIUM payoff — DONE
**Files**: `vendor/univer-revamp/packages/engine-render/src/components/sheets/sheet.render-skeleton.ts:411`

Landed as fork commit `52d85ec78` (perf(engine-render): stop re-walking
the visible span in setStylesCache).

The original incremental-scroll claim in this section was partially
obsolete — the `isIncrementalScroll` branch at line 359-367 already
scopes `styleRanges` to `diffCacheBounds`. The remaining win was a
**redundant inner loop** at line 411-413:

```ts
// Primary — populates bg + border + font for visible cells.
for (let c = visibleStartColumn; c <= visibleEndColumn; c++) {
    this._setStylesCacheForOneCell(r, c, { cacheItem: { bg: true, border: true } });
}
// "Leftward overflow" loop — previously ran [expandStartCol,
// visibleEndColumn), so it re-walked the entire visible span and
// re-called the expensive worksheet.getCell on every visible cell.
for (let c = expandStartCol; c < visibleEndColumn; c++) {
    this._setStylesCacheForOneCell(r, c, { cacheItem: { bg: false, border: false } });
}
```

The cacheItem flag gated `_setBgStylesCache` / `_setBorderStylesCache`
but did NOT gate `_setFontStylesCache`, so each visible cell ran the
expensive helper twice per redraw. Capping the second loop at
`visibleStartColumn` removes the overlap with no rendering change
(font cache values match across the two passes).

**Payoff achieved**: ~30-50% reduction in `setStylesCache` work on
incremental scroll + full redraw at the per-cell level. **Risk**:
verified — all 212 existing engine-render unit tests pass.

### 6. Sparse insert/delete for row + column ops — MEDIUM payoff
**Files**: `vendor/univer/packages/sheets/src/commands/mutations/insert-row-col.mutation.ts:66`,
`vendor/univer/packages/core/src/shared/object-matrix.ts:89-109`

`ObjectMatrix.insertColumns()` / `insertRows()` shift every downstream
cell entry in a loop. On a 10k-row sheet inserting at row 1, all 10k
keys get rewritten. Real sheets are sparse — most of those rows are
empty, but the iteration cost is the same.

**Fix**: track a per-axis "shift offset" map and lazy-remap on access.
Only physically rewrite when offsets accumulate past a threshold (or on
serialize). **Payoff**: medium-to-large (~5-20x faster insert/delete on
large sheets). **Risk**: medium — affects core data model + serialize
path; extensive test coverage required.

### 7. Workbook bootstrap: lazy row/column accumulation — SMALL payoff
**Files**: `vendor/univer/packages/core/src/sheets/sheet-snapshot-utils.ts:47-91`

Loading a workbook materializes the full row/column accumulation arrays
up to the configured size (typically 1000×20), even when actual data is
sparse and confined to A1:E10.

**Fix**: defer `_rowHeightAccumulation` / `_columnWidthAccumulation`
allocation until the first layout pass needs them; allocate only over
ranges that actually carry data or non-default sizing. **Payoff**:
small in time, real in memory on large empty sheets. **Risk**: low.

### 8. Mutation listener fan-out batching — SMALL payoff
**Files**: `vendor/univer/packages/core/src/services/command/command.service.ts:440-456`

After every mutation, the `_collabMutationListeners` forEach runs all
5-10 subscribers (collab bridge, undo stack, dependency tracker, UI).
A 50-cell paste split across 50 single-cell mutations costs 500
listener calls. The `syncOnly: true` flag (line 449) exists but
isn't universally honored.

**Fix**: add a `batched()` wrapper that defers listener calls until the
outer batch resolves. Encourage callers to compound multi-cell ops into
a single mutation (already the pattern in many places — gap is at the
ingress, not the bus). **Payoff**: small (~10-20% latency on bulk
ops). **Risk**: low — listener contracts are well-defined.

---

## High-payoff, HIGH risk (formula engine)

### 9. Formula dirty-range coalescing — LARGE payoff
**Files**: `vendor/univer/packages/sheets-formula/src/controllers/active-dirty.controller.ts`,
`vendor/univer/packages/engine-formula/src/services/calculate-formula.service.ts:154-172`,
`vendor/univer/packages/engine-formula/src/services/dependency-manager.service.ts`

Each mutation marks a dirty range and triggers `getDependencyTree()`.
Adjacent mutations don't coalesce — a 100-cell paste fires 100+
dependency-search queries with overlapping ranges.

**Fix**: batch dirty mutations within a microtask. Merge adjacent /
overlapping dirty rectangles before the dependency search runs.
**Payoff**: large (~5-10x on bulk-edit paths). **Risk**: high —
recalculation correctness depends on the coalesced set being a
superset of the original; missing a dependency means stale results.
Validate exhaustively against the formula test suite.

### 10. Incremental dependency-tree updates — LARGE payoff
**Files**: `vendor/univer/packages/engine-formula/src/services/dependency-manager.service.ts:56`,
`vendor/univer/packages/engine-formula/src/engine/dependency/dependency-tree.ts`

Today, changing one formula can rebuild the dependency tree for every
formula in the workbook. Full rebuild is the conservatively-correct
default but it's expensive at 10k+ formulas.

**Fix**: track the affected subtree only — the edited cell plus its
transitive dependents. Cache untouched branches. **Payoff**: large
(~5-20x on formula edits in big workbooks). **Risk**: high — circular
reference detection must remain valid under incremental updates; this
is the area most likely to introduce subtle correctness bugs. Defer
until coalescing (item 9) is shipped and battle-tested.

---

## Tiny wins (optional)

### 11. Canvas transform batching
`vendor/univer/packages/engine-render/src/components/sheets/spreadsheet.ts:315, 364` —
buffer `translateWithPrecision` + `setTransform` to flush once per
frame instead of per op. Rarely measurable; <1% impact.

### 12. Lazy initialization of optional services
Several render / docs services boot on workbook open whether or not the
feature is reached. Audit and defer to first-use. Bootstrap-time
savings; doesn't move the steady-state needle.

---

## Sequencing recommendation

Ship in groups, validate between groups:

1. **Group A (Quick wins, ~1 week)**: items 1-4. Pure perf, low blast
   radius, no model changes.
2. **Group B (Render pipeline, ~1 week)**: items 5, 7. Cleaner
   incremental scroll + lazy bootstrap.
3. **Group C (Data model, ~2 weeks)**: items 6, 8. Sparse shifts +
   listener batching — requires regression coverage of insert/delete +
   mutation ordering.
4. **Group D (Formula engine, multi-week)**: items 9, 10. Coalescing
   first (lower risk of the two), then incremental tree. Each behind a
   feature flag until the formula test suite passes twice over.

Expected end-state: **2-5× on bulk mutations**, **20-40% on scrolling**,
**10-20% on per-edit latency**.

---

## Out of scope (for now)

- Worker offload of layout: the formula engine already runs in a
  worker via `UniverRPCWorkerThreadPlugin`. Moving layout there too is
  a bigger rewrite than this plan covers.
- WASM hot loops (border rendering, dependency walk): potentially
  large win but adds a build-toolchain dependency. Re-evaluate after
  Group D ships.
- Pro-only features: charts/pivots/print live in upstream's commercial
  layer; this plan only touches OSS.
