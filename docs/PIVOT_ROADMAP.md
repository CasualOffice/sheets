# PivotTable roadmap

State of the in-house pivot (built on OSS, no Univer Pro) and the design for the
remaining work. The engine is a pure compute (`apps/web/src/pivots/compute.ts`)
feeding a single `setValues` write (`apply.ts`); models live on
`IWorkbookData.resources` and round-trip via the hidden sidecar sheet.

## Shipped

| Capability                                                       | PR         |
| ---------------------------------------------------------------- | ---------- |
| Single/multi **row** fields (compact layout, indented subtotals) | P1         |
| Single **column** field → cross-tab / matrix                     | P2         |
| Filters                                                          | P1         |
| Aggregations: Sum / Count / Average / Min / Max                  | P0         |
| **Multiple value fields** (one output column each)               | #211       |
| **Distinct Count**                                               | #212       |
| **Show Values As → % of Grand Total** (row + cross-tab)          | #214, #218 |
| **Show Values As → % of Column / Row Total**                     | #223, #224 |
| **Date grouping** of the primary row field (Year/Quarter/Month)  | #215       |
| **Nested (multiple) column fields** → tuple cross-tab            | #220, #221 |
| Drill-down (double-click a value → source rows)                  | P2         |
| **PivotTable Fields pane** — reconfigure zones + live re-apply   | #229       |

All of the above are pure-compute + dialog/panel increments with unit + e2e
coverage. The nested-columns compute rewrite (designed below) shipped in
#220/#221; the field pane (the last big UI subsystem) shipped its first slice in
#229. Remaining work is the drag-and-drop layer on the pane plus the live,
refreshable Excel-native pivot object.

## Design record: nested column fields (shipped #220/#221)

_Kept as the design rationale for the nested-column compute; implemented as
described._

**Why it's not a small increment.** `computeMatrix` is built around a _single_
column field: `colFieldCol = model.cols[0]`, `colKeys` is a flat list of that
field's distinct values, each value row fans out one block per key, and
`colMeta` (which **drill-down** relies on to map a clicked column back to its
key) is indexed per single key. Nesting changes all four:

- **Column keys become tuples** — the cartesian product of each column field's
  values that actually occur (e.g. `Quarter × Month` → `Q1·Jan, Q1·Feb, …`),
  sorted lexicographically per level.
- **Header becomes N+1 rows** — one spanning row per column field (outer key
  repeated across its sub-columns, rest blanked — same no-merge convention we
  use today for the multi-value sub-header) plus the value-field sub-header.
- **`valueCellsFor` slices by the full tuple** — `records.filter` matching every
  column field, not just one.
- **`colMeta` carries the key path** — `{ kind:'value'; colKeys: string[];
valueIndex }` so drill-down keeps working; this is the load-bearing change to
  verify, since drill-down decodes `colMeta` to rebuild the clicked subset.

**Proposed shape.** Generalise the single-field matrix to an ordered
`colFields = model.cols.map(c => c.column)`:

1. Build `colKeyTuples`: distinct `colFields.map(c => keyOf(rec[c]))` over the
   filtered records, sorted by level (outer first). Cap the tuple count (e.g.
   2,048) and `log()`/surface a notice if exceeded — a 3-deep nest over
   high-cardinality fields explodes the column count and would freeze the write.
2. Headers: emit `colFields.length` spanning rows + (if multi-value) the value
   sub-header; grand-total block unchanged.
3. `valueCellsFor(records)`: for each tuple, slice by all fields, aggregate each
   value; then the across-all grand-total block (unchanged).
4. `colMeta`: `{ kind:'value', colKeys: string[], valueIndex }` /
   `{ kind:'grand-total', valueIndex }`; update drill-down's column decoder to
   read `colKeys` (was `colKey`).
5. Show-Values-As post-pass already keys off `colMeta[].valueIndex` — works
   unchanged once `colMeta` carries tuples.

**Dialog.** Replace the single "Column field" select with an add/remove list
(mirroring the value-field list from #211); default one, cap at 2–3 levels.

**Risk / test plan.** Drill-down is the regression surface. Cover: 2-level
nest single-value; 2-level nest multi-value (header span correctness); a tuple
that's absent in the data (no empty column); drill-down on a nested column;
Show-Values-As % over a nested matrix; the tuple-cap notice. Keep the existing
single-column-field specs green (a single `cols` entry must produce identical
output to today).

Estimate: ~1 focused PR for compute + colMeta + drill-down, ~1 for the dialog +
e2e. Do compute first behind the existing single-field dialog (model already
allows `cols: PivotFieldRef[]`), then ship the multi-select dialog.

## Field pane (slice 1 shipped #229) + live object

**PivotTable Fields pane — slice 1 (#229, shipped).** A rail panel
(`PivotFieldsPanel`, like Charts/Tables) that reflects a pivot across the four
Excel zones (Filters / Columns / Rows / Values) and re-applies it **live** as
you reconfigure — no delete-and-reinsert. Field assignment uses a click "+"
menu (Excel's right-click "Add to Row Labels / …" affordance), with per-chip
remove and reorder, Values aggregation + Show-Values-As editing, and Rows
date-grouping. The model ops are pure + unit-tested (`fields-model.ts`); the
wiring is e2e-tested (`pivot-fields-panel.spec.ts`). Auto-opens on insert. Also
fixed a latent bug: `resources.ts` `VALID_AGGS` omitted `distinctCount`,
silently dropping a saved Distinct-Count pivot on reload.

**Slice 2 — report-filter value selection (shipped).** The Filters zone now
carries a per-value checklist (expand a filter chip → check/uncheck values,
Select-all / Clear) that actually narrows the source records and re-applies the
pivot live. Toggle maths is pure + unit-tested (`toggleFilterValue` /
`setFilterValues` in `fields-model.ts`); wiring is e2e-tested
(`pivot-fields-report-filter.spec.ts`).

**Slice 3a — auto-follow the active selection (shipped).** Clicking into a
pivot's output switches the pane to that pivot (via `findPivotAtCell`),
Excel-style; clicking away leaves it on the last one. e2e:
`pivot-fields-autofollow.spec.ts`.

**Slice 3b (next) — drag-and-drop between zones.** Layered on the same
`fields-model` ops (the drop reducer is just `addFieldToZone` /
`moveWithinZone`). Click-to-assign stays as the reliable / accessible path;
DnD is the polish layer. Note: HTML5 DnD is flaky to e2e in Playwright, so the
plan is unit-tested drop logic + a tolerant drag spec rather than relying on a
brittle native-drag test.

**Live, refreshable pivot object.** Today a pivot is materialised cells + a
model resource; "refresh" re-runs compute. A true live object (auto-refresh on
source change, Excel-native `xl/pivotTables` on export) is a larger effort and
overlaps the raw-OOXML pivot passthrough already in the xlsx bridge.

## Deferred (Excel parity, lower demand)

Running totals (% of … running / rank); calculated fields; value filters /
top-N; manual sort & grouping of non-date fields.
