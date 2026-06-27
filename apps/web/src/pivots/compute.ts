import type { PivotAggregation, PivotModel } from './types';
import { PIVOT_AGG_LABELS } from './types';

/**
 * Pure compute step — turn raw source records into the laid-out cell
 * grid that gets written into the workbook. No FUniver dependency
 * here so it's trivially testable and so the same routine can run
 * server-side later (e.g. for an "export pivot to CSV" path).
 *
 * Scope as of P2:
 *   - 1..N row fields (multi-row uses Excel's compact layout — see
 *     {@link computePivot} below).
 *   - Optional single column field → cross-tab / matrix layout (see
 *     {@link computeMatrix} below). When present, the value field(s)
 *     fan out across one block of columns per distinct column-field
 *     key, with a "Grand Total" block on the right and a Grand Total
 *     row at the bottom.
 *   - One or more value fields, each gets its own column (no column
 *     field) or its own column *within each* column-key block (with a
 *     column field).
 *   - Filters applied before bucketing (P1).
 *
 * Matrix layout (rows = [Region], cols = [Quarter], values = [Sum Sales]):
 *
 *     [ Region   | Q1  | Q2  | Grand Total ]
 *     [ North    | 100 | 120 |         220 ]
 *     [ South    |  80 |  95 |         175 ]
 *     [ Grand T. | 180 | 215 |         395 ]
 *
 * Single-row output (unchanged from P0):
 *
 *     [ row-field | Sum of Sales | Avg of Sales ]
 *     [ North     |          220 |          110 ]
 *     [ South     |          175 |         87.5 ]
 *     [ Grand T.  |          395 |         98.7 ]
 *
 * Multi-row compact layout (rows = [Region, Product], values = [Sum]):
 *
 *     [ Region    | Sum of Sales ]
 *     [ East      |          300 ]   ← outer subtotal on the label row
 *     [   A       |          100 ]   ← inner leaf, indent depth 1
 *     [   B       |          200 ]
 *     [ West      |          150 ]
 *     [   A       |          150 ]
 *     [ Grand T.  |          450 ]
 *
 * Indentation uses leading spaces in the label string (`'  '.repeat
 * (depth)`) — Univer's IStyleData has no first-class indent property,
 * and spaces round-trip losslessly through xlsx.
 */
export type PivotCell = string | number | null;
export type PivotGrid = PivotCell[][];

/** Metadata for each row of the output grid — lets drill-down map a
 *  clicked cell back to the composite key path that produced it
 *  without re-running the bucketing walk. */
export type PivotRowMeta =
  | { kind: 'header' }
  | { kind: 'subtotal'; keyPath: string[]; depth: number }
  | { kind: 'leaf'; keyPath: string[]; depth: number }
  | { kind: 'grand-total' };

/** Metadata for each COLUMN of a matrix (cross-tab) output grid — lets
 *  drill-down map a clicked column back to its column-field key. Only
 *  populated when the pivot has a column field; non-matrix pivots leave
 *  `colMeta` undefined.
 *
 *   - `label`     → the leading label column (row-field names).
 *   - `value`     → a value cell scoped to a single column-field key.
 *   - `grand-total` → a value cell aggregating across all column keys
 *     (the right-hand "Grand Total" block).
 */
export type PivotColMeta =
  | { kind: 'label' }
  | { kind: 'value'; colKey: string; valueIndex: number }
  | { kind: 'grand-total'; valueIndex: number };

export type PivotComputeResult = {
  grid: PivotGrid;
  rowMeta: PivotRowMeta[];
  /** Present only for matrix (column-field) pivots. Index-aligned with
   *  the columns of every grid row. */
  colMeta?: PivotColMeta[];
};

/**
 * Raw records read from the workbook source range. Row 0 is the
 * header row, rows 1.. are data. Cell values come in pre-coerced
 * by the caller (`apply.ts` reads them via `getValue()` which already
 * returns the right primitive type).
 */
export type SourceMatrix = {
  headers: string[];
  records: PivotCell[][];
};

export function computePivot(source: SourceMatrix, model: PivotModel): PivotComputeResult {
  if (model.values.length === 0) {
    return { grid: [], rowMeta: [] };
  }

  // P2 — a column field switches the layout to a cross-tab / matrix.
  // The value field(s) fan out horizontally across one block per
  // distinct column key. Single-row + multi-row both supported. We
  // only honour the first column field; nested column fields are a
  // future follow-up (the model array shape allows it without a bump).
  if ((model.cols?.length ?? 0) > 0) {
    return computeMatrix(source, model);
  }

  // P1 — apply filter fields BEFORE bucketing. Each filter restricts
  // records to those whose value in `column` is one of `allowedValues`
  // (string-compared). Empty allowedValues excludes everything; the
  // dialog UI never produces that shape (uses remove-filter), but
  // compute defends against it.
  const filters = model.filters ?? [];
  const passesFilters = (rec: PivotCell[]): boolean => {
    for (const f of filters) {
      const allowed = new Set(f.allowedValues);
      const v = rec[f.column];
      const key = v == null ? '' : String(v);
      if (!allowed.has(key)) return false;
    }
    return true;
  };
  const filteredRecords =
    filters.length > 0 ? source.records.filter(passesFilters) : source.records;

  // Header row — column 0 is the outermost row field name (compact
  // layout uses one label column shared across all row-field levels);
  // subsequent columns are the value-field headers in model order.
  const rowFieldCols = model.rows.map((r) => r.column);
  const hasRowField = rowFieldCols.length > 0;
  const header: PivotCell[] = [];
  if (hasRowField) {
    header.push(source.headers[rowFieldCols[0]] ?? '');
  }
  for (const v of model.values) {
    header.push(`${PIVOT_AGG_LABELS[v.agg]} of ${source.headers[v.column] ?? ''}`);
  }

  const grid: PivotGrid = [header];
  const rowMeta: PivotRowMeta[] = [{ kind: 'header' }];

  if (!hasRowField) {
    // No row field — Grand Total only.
    const total: PivotCell[] = [];
    for (const v of model.values) {
      total.push(
        aggregate(
          filteredRecords.map((r) => r[v.column]),
          v.agg,
        ),
      );
    }
    grid.push(total);
    rowMeta.push({ kind: 'grand-total' });
    return { grid, rowMeta };
  }

  // Build the nested bucket tree. Each level keys by the value of the
  // corresponding row field; leaves hold the contributing record list.
  const root = buildTree(filteredRecords, rowFieldCols);

  // Walk the tree in compact-layout order. For each row-field level we
  // emit either a subtotal row (intermediate levels) or a leaf row
  // (innermost level), then recurse into children.
  const walk = (node: TreeNode, keyPath: string[], depth: number): void => {
    // Sort keys ascending — Excel default sort order.
    const keys = [...node.children.keys()].sort((a, b) =>
      a.localeCompare(b, undefined, { numeric: true }),
    );
    const isInnermost = depth === rowFieldCols.length - 1;
    for (const key of keys) {
      const child = node.children.get(key)!;
      const path = [...keyPath, key];
      const cells: PivotCell[] = [`${'  '.repeat(depth)}${key}`];
      for (const v of model.values) {
        cells.push(
          aggregate(
            child.records.map((r) => r[v.column]),
            v.agg,
          ),
        );
      }
      grid.push(cells);
      rowMeta.push(
        isInnermost
          ? { kind: 'leaf', keyPath: path, depth }
          : { kind: 'subtotal', keyPath: path, depth },
      );
      if (!isInnermost) walk(child, path, depth + 1);
    }
  };
  walk(root, [], 0);

  // Grand Total — aggregates ALL filtered records (regardless of row
  // bucketing depth) so the total always matches the visible records.
  const total: PivotCell[] = ['Grand Total'];
  for (const v of model.values) {
    total.push(
      aggregate(
        filteredRecords.map((r) => r[v.column]),
        v.agg,
      ),
    );
  }
  grid.push(total);
  rowMeta.push({ kind: 'grand-total' });

  return { grid, rowMeta };
}

/**
 * Cross-tab / matrix layout. Active when `model.cols` is non-empty.
 *
 * Columns fan out by the distinct values of the (first) column field;
 * within each column-key block there is one column per value field
 * (one block === one column for the common single-value case). A
 * right-hand "Grand Total" block aggregates across all column keys,
 * and a bottom Grand Total row aggregates down each column.
 *
 * Rows reuse the same compact bucket walk as {@link computePivot} so
 * multi-row matrices indent inner keys exactly like the non-matrix
 * path. The difference is purely horizontal: each row's value cells
 * are sliced per column key instead of being a single total.
 */
function computeMatrix(source: SourceMatrix, model: PivotModel): PivotComputeResult {
  const filters = model.filters ?? [];
  const passesFilters = (rec: PivotCell[]): boolean => {
    for (const f of filters) {
      const allowed = new Set(f.allowedValues);
      const v = rec[f.column];
      const key = v == null ? '' : String(v);
      if (!allowed.has(key)) return false;
    }
    return true;
  };
  const filteredRecords =
    filters.length > 0 ? source.records.filter(passesFilters) : source.records;

  const colFieldCol = model.cols[0].column;
  const rowFieldCols = model.rows.map((r) => r.column);
  const hasRowField = rowFieldCols.length > 0;
  const values = model.values;
  const multiValue = values.length > 1;

  // Distinct column-field keys, sorted Excel-default (numeric-aware
  // ascending). Computed from the filtered records so a filter that
  // removes every record for a column key drops that column entirely
  // (matches Excel — empty columns don't appear).
  const colKeys = [...new Set(filteredRecords.map((r) => keyOf(r[colFieldCol])))].sort((a, b) =>
    a.localeCompare(b, undefined, { numeric: true }),
  );

  // ---- Header rows -------------------------------------------------
  // Single value field → one header row: [label, ...colKeys, Grand T].
  // Multiple value fields → two header rows: the column-key spanning
  // the value-field sub-headers beneath it. The spanning row repeats
  // the key in its first sub-cell and blanks the rest (Univer has no
  // cell-merge in a plain setRangeValues write; the label still reads
  // left-to-right and round-trips through xlsx).
  const labelHeader = hasRowField ? (source.headers[rowFieldCols[0]] ?? '') : '';

  const colMeta: PivotColMeta[] = [{ kind: 'label' }];
  for (const key of colKeys) {
    for (let vi = 0; vi < values.length; vi += 1) {
      colMeta.push({ kind: 'value', colKey: key, valueIndex: vi });
    }
  }
  for (let vi = 0; vi < values.length; vi += 1) {
    colMeta.push({ kind: 'grand-total', valueIndex: vi });
  }

  const valueLabel = (vi: number): string =>
    `${PIVOT_AGG_LABELS[values[vi].agg]} of ${source.headers[values[vi].column] ?? ''}`;

  const grid: PivotGrid = [];
  const rowMeta: PivotRowMeta[] = [];

  if (multiValue) {
    // Top header row: column-key spans.
    const top: PivotCell[] = [labelHeader];
    for (const key of colKeys) {
      top.push(key === '' ? '(blank)' : key);
      for (let i = 1; i < values.length; i += 1) top.push('');
    }
    top.push('Grand Total');
    for (let i = 1; i < values.length; i += 1) top.push('');
    grid.push(top);
    rowMeta.push({ kind: 'header' });
    // Sub-header row: value-field labels under each key + grand total.
    const sub: PivotCell[] = [''];
    for (let c = 0; c < colKeys.length; c += 1) {
      for (let vi = 0; vi < values.length; vi += 1) sub.push(valueLabel(vi));
    }
    for (let vi = 0; vi < values.length; vi += 1) sub.push(valueLabel(vi));
    grid.push(sub);
    rowMeta.push({ kind: 'header' });
  } else {
    const header: PivotCell[] = [labelHeader];
    for (const key of colKeys) header.push(key === '' ? '(blank)' : key);
    header.push('Grand Total');
    grid.push(header);
    rowMeta.push({ kind: 'header' });
  }

  // colMeta is only meaningful with a single header row. For the
  // multi-value two-row header the second header row shares the same
  // column structure, so the same colMeta still index-aligns to the
  // VALUE rows below — which is all drill-down needs.

  // ---- Value rows --------------------------------------------------
  // Compute one row of value cells for a given record subset: for each
  // column key, slice the subset to records matching that key, then
  // aggregate each value field; finish with the across-all grand-total
  // block.
  const valueCellsFor = (records: PivotCell[][]): PivotCell[] => {
    const cells: PivotCell[] = [];
    for (const key of colKeys) {
      const slice = records.filter((rec) => keyOf(rec[colFieldCol]) === key);
      for (const v of values) {
        cells.push(
          aggregate(
            slice.map((r) => r[v.column]),
            v.agg,
          ),
        );
      }
    }
    // Grand-total block — aggregate across every column key (i.e. the
    // whole row subset, ignoring the column split).
    for (const v of values) {
      cells.push(
        aggregate(
          records.map((r) => r[v.column]),
          v.agg,
        ),
      );
    }
    return cells;
  };

  if (!hasRowField) {
    // No row field — a single Grand Total row carrying the column split.
    grid.push(['Grand Total', ...valueCellsFor(filteredRecords)]);
    rowMeta.push({ kind: 'grand-total' });
    return { grid, rowMeta, colMeta };
  }

  const root = buildTree(filteredRecords, rowFieldCols);
  const walk = (node: TreeNode, keyPath: string[], depth: number): void => {
    const keys = [...node.children.keys()].sort((a, b) =>
      a.localeCompare(b, undefined, { numeric: true }),
    );
    const isInnermost = depth === rowFieldCols.length - 1;
    for (const key of keys) {
      const child = node.children.get(key)!;
      const path = [...keyPath, key];
      grid.push([`${'  '.repeat(depth)}${key}`, ...valueCellsFor(child.records)]);
      rowMeta.push(
        isInnermost
          ? { kind: 'leaf', keyPath: path, depth }
          : { kind: 'subtotal', keyPath: path, depth },
      );
      if (!isInnermost) walk(child, path, depth + 1);
    }
  };
  walk(root, [], 0);

  // Bottom Grand Total row — column totals + the overall total.
  grid.push(['Grand Total', ...valueCellsFor(filteredRecords)]);
  rowMeta.push({ kind: 'grand-total' });

  return { grid, rowMeta, colMeta };
}

/** Coerce a cell value to its string bucket key (null/empty → ''). */
function keyOf(v: PivotCell): string {
  return v == null ? '' : String(v);
}

type TreeNode = {
  records: PivotCell[][];
  children: Map<string, TreeNode>;
};

function buildTree(records: PivotCell[][], rowFieldCols: number[]): TreeNode {
  const root: TreeNode = { records: [], children: new Map() };
  for (const rec of records) {
    let node = root;
    node.records.push(rec);
    for (const col of rowFieldCols) {
      const key = rec[col] == null ? '' : String(rec[col]);
      let child = node.children.get(key);
      if (!child) {
        child = { records: [], children: new Map() };
        node.children.set(key, child);
      }
      child.records.push(rec);
      node = child;
    }
  }
  return root;
}

function aggregate(values: PivotCell[], agg: PivotAggregation): PivotCell {
  const nums: number[] = [];
  let nonNull = 0;
  for (const v of values) {
    if (v == null || v === '') continue;
    nonNull++;
    const n = typeof v === 'number' ? v : Number(v);
    if (Number.isFinite(n)) nums.push(n);
  }
  switch (agg) {
    case 'count':
      return nonNull;
    case 'sum':
      return nums.length === 0 ? 0 : nums.reduce((a, b) => a + b, 0);
    case 'average':
      return nums.length === 0 ? null : nums.reduce((a, b) => a + b, 0) / nums.length;
    case 'min':
      return nums.length === 0 ? null : Math.min(...nums);
    case 'max':
      return nums.length === 0 ? null : Math.max(...nums);
    case 'distinctCount': {
      // Count distinct non-empty values (compared as strings, like Excel's
      // "Distinct Count"). Empty/blank cells are ignored.
      const seen = new Set<string>();
      for (const v of values) {
        if (v == null || v === '') continue;
        seen.add(String(v));
      }
      return seen.size;
    }
  }
}

/** Used by the panel to render a friendly auto-name like
 *  "Sum of Sales by Region" so a freshly inserted pivot has a label
 *  that explains itself. */
export function defaultPivotTitle(source: SourceMatrix, model: PivotModel): string {
  const value = model.values[0];
  if (!value) return 'PivotTable';
  const valuePart = `${PIVOT_AGG_LABELS[value.agg]} of ${source.headers[value.column] ?? 'value'}`;
  const rowFields = model.rows.map((r) => source.headers[r.column] ?? 'group');
  const colFields = (model.cols ?? []).map((c) => source.headers[c.column] ?? 'group');
  const byPart = rowFields.length === 0 ? '' : ` by ${rowFields.join(' / ')}`;
  const acrossPart = colFields.length === 0 ? '' : ` across ${colFields.join(' / ')}`;
  if (!byPart && !acrossPart) return valuePart;
  return `${valuePart}${byPart}${acrossPart}`;
}
