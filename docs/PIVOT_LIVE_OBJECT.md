# Live / refreshable pivot object — design

Design pass for the last Excel-parity gap in the pivot stack. The Fields pane
(`docs/PIVOT_ROADMAP.md`) made pivots interactively reconfigurable; what's left
is making an **in-app-created** pivot a _real_ pivot to the outside world —
recognised as a PivotTable when the file is opened in Excel — and tightening the
refresh lifecycle. This doc states the approach and the decisions that need a
greenlight before coding.

## Current state (three layers)

1. **Materialised cells.** `apply.ts` computes the grid and writes it with one
   `setRangeValues`. This is what the user sees and what survives every path.
2. **`PivotModel` resource + Fields pane.** The definition lives on
   `IWorkbookData.resources` (`__casual_sheets_pivots__`), round-trips through
   xlsx via the hidden resources sheet, and drives reconfiguration + manual
   `Refresh PivotTables`.
3. **`pivot-passthrough.ts`.** For pivots that arrive in an **imported** xlsx,
   the raw OOXML parts (`xl/pivotCaches/*`, `xl/pivotTables/*`, their rels,
   `[Content_Types]`, the `workbook.xml` `<pivotCaches>` element) are captured
   at parse (`capturePivotsFromBuffer`) and re-injected at export
   (`applyPivotsToZip`). Excel re-recognises them; we never edit their contents.

## The gap

- **In-app-created pivots export as flat values.** A pivot the user builds here
  has layers 1–2 but **no OOXML pivot parts**. Saved and reopened in Excel, it's
  just a block of numbers — no field list, no refresh, no "PivotTable Tools".
  This is the headline gap.
- **Refresh lifecycle is manual-only.** We have a global `Refresh PivotTables`
  menu item. Excel additionally offers per-pivot refresh and a per-pivot
  "Refresh data when opening the file" option.

## Goal & scope

Make an in-app pivot indistinguishable from an Excel-authored one on
round-trip, and round out refresh. Explicitly **out of scope**: pivot charts,
OLAP/data-model pivots, calculated fields/items (tracked separately in the
roadmap's Deferred list).

## Workstream A — native export of in-app pivots (the big one)

Synthesise the OOXML parts from `PivotModel` at export and inject them with the
**machinery `pivot-passthrough.ts` already proves out** (rel renumbering,
`[Content_Types]` overrides, `workbook.xml` `<pivotCaches>` surgery, per-sheet
`.rels`). We are not inventing the injection layer — only the **part
generation**.

Parts to generate per in-app pivot:

| Part                                         | Generated from `PivotModel`                                                                                                                                                                             |
| -------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `xl/pivotCaches/pivotCacheDefinition{N}.xml` | `source` range → `<cacheSource>`; header row → `<cacheFields>` + shared items per field                                                                                                                 |
| `xl/pivotCaches/pivotCacheRecords{N}.xml`    | the source records, encoded as shared-item indices / inline values                                                                                                                                      |
| `xl/pivotTables/pivotTable{N}.xml`           | `rows`/`cols`/`values`/`filters` → `<rowFields>`/`<colFields>`/`<dataFields>`/`<pageFields>`; `target` → `<location ref>`; `grouping` → field group; `agg`/`showAs` → `<dataField subtotal/showDataAs>` |
| each part's `.rels`                          | fixed templates (pivotTable→cacheDefinition, cacheDefinition→cacheRecords)                                                                                                                              |

Then the shared injection steps (already implemented for passthrough): remap
rIds into `workbook.xml.rels` + the target sheet's `.rels`, splice
`<pivotCaches>` into `workbook.xml`, add `Override` entries to
`[Content_Types].xml`.

**Coexistence with the passthrough.** Tag each pivot's origin so we never
double-emit: imported pivots flow through `applyPivotsToZip` (verbatim);
in-app pivots flow through the new generator. `PivotModel` gets an
`origin: 'app' | 'imported'` discriminant (default `'app'`; the importer sets
`'imported'` and links to the captured raw parts). Numbering (`{N}`) is
allocated across both sets so rIds and cache ids don't collide.

**Cache vs. live data.** `pivotCacheRecords` is a _snapshot_. Excel refreshes
it from `cacheSource` on demand. We write the snapshot to match our
materialised cells at export time; set `<pivotCacheDefinition refreshOnLoad="1">`
(see Workstream B) so Excel recomputes from the source on open and the two can't
drift.

**Risks / unknowns to spike first:**

- **Shared-items fidelity.** Numbers vs. strings vs. dates each have distinct
  `<sharedItems>` encodings (`containsNumber`, `containsDate`, …). Getting these
  wrong makes Excel repair the file. Spike with a 2-field numeric+string pivot
  and diff against an Excel-authored equivalent.
- **`showDataAs` mapping.** Our `PivotShowAs` (% of grand/row/column total) maps
  to `<dataField showDataAs=…>` — verify each enum value.
- **Date grouping** maps to `<fieldGroup>` with a range group; non-trivial,
  could be phase 2 of this workstream.
- **Repair-on-open is binary feedback.** Excel either accepts the file silently
  or shows "we found a problem." Build a fixture harness that opens generated
  files in a headless validator (or at minimum, manual Excel checks gated in the
  PR description) — unit tests on the XML strings are necessary but not
  sufficient.

## Workstream B — refresh lifecycle (small, Excel-aligned)

- **Refresh on open.** Recompute every pivot once after a workbook loads (hook
  the existing load/`revision`-bump path the Fields pane already uses). This is
  Excel's per-pivot "Refresh data when opening the file"; making it the default
  guarantees an opened file shows correct numbers and keeps the exported cache
  honest.
- **Per-pivot refresh.** A refresh affordance on the pivot (Fields pane header /
  context) in addition to the global menu item.
- **Decision: no auto-refresh-on-edit.** Excel does _not_ live-update pivots as
  you edit source cells (you click Refresh). We match that — auto-on-edit would
  be a Google-Sheets behaviour, off-bar, and risks edit-loop/perf issues.
  Refresh stays on-demand and on-open only.

## Phasing

1. **B first** (refresh-on-open + per-pivot refresh) — small, independent, ships
   value immediately and makes A's exported cache trustworthy.
2. **A.1 spike** — generate the three parts for one simple pivot (single row
   field, single Sum value, no columns/filters/grouping); inject; confirm Excel
   opens it clean. This de-risks the whole workstream.
3. **A.2** — columns, multiple values, filters, `showDataAs`.
4. **A.3** — date grouping → `<fieldGroup>`.

Each of A.1–A.3 is its own PR with XML-string unit tests + a round-trip e2e
(generate → re-parse with `capturePivotsFromBuffer` → assert structure) and a
manual Excel-open check noted in the PR.

## Decisions needing a greenlight

1. **Default `refreshOnLoad="1"`** on generated caches (recommended — keeps
   cache ⇄ source honest). Confirm.
2. **`origin` discriminant on `PivotModel`** as the import/in-app split.
   Confirm the field + that imported pivots stay verbatim-passthrough (we don't
   try to regenerate them from a reverse-engineered model).
3. **Validation bar for "Excel opens it clean."** Is a manual Excel check per PR
   acceptable, or do we want a headless OOXML validator in CI (larger setup)?

Until these are settled, Workstream B is safe to start independently.
