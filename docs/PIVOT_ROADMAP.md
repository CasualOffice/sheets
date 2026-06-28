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
| **Date grouping** of the primary row field (Year/Quarter/Month)  | #215       |
| Drill-down (double-click a value → source rows)                  | P2         |

All of the above are pure-compute + dialog increments with unit + e2e coverage.
That seam is now exhausted for cleanly-bounded work — the remaining items are
either a compute **rewrite** (nested columns) or a new **UI subsystem** (field
pane, live object). They need a design pass first; this doc is that pass.

## Next: multiple (nested) column fields

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

## Later: field pane + live object (out of scope until the above lands)

- **Drag-and-drop PivotTable Fields pane** — the Excel side panel (drag fields
  between Rows/Columns/Values/Filters). A real UI subsystem (a panel in the
  rail, like Charts/Tables panels) on top of the existing model. Multi-day;
  design separately.
- **Live, refreshable pivot object** — today a pivot is materialised cells +
  a model resource; "refresh" re-runs compute. A true live object (auto-refresh
  on source change, Excel-native `xl/pivotTables` on export) is a larger effort
  and overlaps the raw-OOXML pivot passthrough already in the xlsx bridge.

## Deferred (Excel parity, lower demand)

Show Values As → % of Column/Row Total + running totals; calculated fields;
value filters / top-N; manual sort & grouping of non-date fields.
