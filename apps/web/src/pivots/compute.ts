import type {
  DateGrouping,
  PivotAggregation,
  PivotFieldRef,
  PivotModel,
  PivotValueField,
} from './types';
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
  | { kind: 'value'; colKeys: string[]; valueIndex: number }
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
  const root = buildTree(filteredRecords, model.rows);

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

  // "Show Values As → % of Grand Total" per value field. The grand-total row
  // (just pushed) holds each value column's denominator; value columns start
  // at index 1 (column 0 is the shared row-field label).
  applyShowAsPercent(grid, model.values, 1);

  return { grid, rowMeta };
}

/**
 * Rewrite each value column flagged `showAs: 'pctOfGrandTotal'` as a percentage
 * of that column's grand total (the last grid row). Mutates `grid` in place;
 * the header row is left untouched and the grand-total cell becomes 100.0%.
 */
function applyShowAsPercent(
  grid: PivotGrid,
  values: PivotValueField[],
  valueColStart: number,
): void {
  if (grid.length < 2) return;
  const lastRow = grid.length - 1;
  values.forEach((v, vi) => {
    // In the row-only layout a value column's "column total" is its grand
    // total, so both percentage modes share the same denominator.
    if (v.showAs !== 'pctOfGrandTotal' && v.showAs !== 'pctOfColumnTotal') return;
    const col = valueColStart + vi;
    const denom = Number(grid[lastRow][col]);
    for (let r = 1; r < grid.length; r++) {
      const raw = grid[r][col];
      const n = typeof raw === 'number' ? raw : Number(raw);
      grid[r][col] = denom && Number.isFinite(n) ? `${((n / denom) * 100).toFixed(1)}%` : '0.0%';
    }
  });
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

  const colFields = model.cols.map((c) => c.column);
  const rowFieldCols = model.rows.map((r) => r.column);
  const hasRowField = rowFieldCols.length > 0;
  const values = model.values;
  const multiValue = values.length > 1;

  const tupleOf = (rec: PivotCell[]): string[] => colFields.map((c) => keyOf(rec[c]));

  // Distinct column-key TUPLES present in the data (one entry per column field,
  // outer-first), sorted level-by-level (numeric-aware). A single column field
  // reduces to 1-element tuples → identical output to the pre-nesting path. A
  // filter that empties a tuple drops that column (Excel — no empty columns).
  const seenTuples = new Map<string, string[]>();
  for (const rec of filteredRecords) {
    const t = tupleOf(rec);
    seenTuples.set(t.join('\u0000'), t);
  }
  let tuples = [...seenTuples.values()].sort((a, b) => {
    for (let i = 0; i < a.length; i += 1) {
      const c = a[i].localeCompare(b[i], undefined, { numeric: true });
      if (c !== 0) return c;
    }
    return 0;
  });
  // A high-cardinality nest can explode the column count (and the single
  // setValues write). Truncate + flag; the dialog also caps the depth.
  const TUPLE_CAP = 2048;
  const tuplesCapped = tuples.length > TUPLE_CAP;
  if (tuplesCapped) tuples = tuples.slice(0, TUPLE_CAP);

  const labelHeader = hasRowField ? (source.headers[rowFieldCols[0]] ?? '') : '';
  const valueLabel = (vi: number): string =>
    `${PIVOT_AGG_LABELS[values[vi].agg]} of ${source.headers[values[vi].column] ?? ''}`;

  // colMeta — label, then one entry per (tuple, value) value cell, then the
  // grand-total block (one per value field). Drill-down decodes `colKeys`.
  const colMeta: PivotColMeta[] = [{ kind: 'label' }];
  for (const t of tuples) {
    for (let vi = 0; vi < values.length; vi += 1) {
      colMeta.push({ kind: 'value', colKeys: t, valueIndex: vi });
    }
  }
  for (let vi = 0; vi < values.length; vi += 1) {
    colMeta.push({ kind: 'grand-total', valueIndex: vi });
  }

  const grid: PivotGrid = [];
  const rowMeta: PivotRowMeta[] = [];

  // ---- Header rows -------------------------------------------------
  // One spanning row per column-field level (outer first), then — when there
  // are multiple value fields — a value-field sub-header row. A key is shown at
  // the start of its span and blanked across the rest (Univer has no cell merge
  // in setValues; labels still read left-to-right and round-trip through xlsx).
  for (let level = 0; level < colFields.length; level += 1) {
    const row: PivotCell[] = [level === 0 ? labelHeader : ''];
    let prevPrefix: string | null = null;
    for (const t of tuples) {
      const prefix = t.slice(0, level + 1).join('\u0000');
      const show = prefix !== prevPrefix;
      prevPrefix = prefix;
      const label = t[level] === '' ? '(blank)' : t[level];
      for (let vi = 0; vi < values.length; vi += 1) {
        row.push(show && vi === 0 ? label : '');
      }
    }
    // Grand-total block — "Grand Total" on the first level row, blank below.
    for (let vi = 0; vi < values.length; vi += 1) {
      row.push(level === 0 && vi === 0 ? 'Grand Total' : '');
    }
    grid.push(row);
    rowMeta.push({ kind: 'header' });
  }
  if (multiValue) {
    const sub: PivotCell[] = [''];
    for (let i = 0; i < tuples.length; i += 1) {
      for (let vi = 0; vi < values.length; vi += 1) sub.push(valueLabel(vi));
    }
    for (let vi = 0; vi < values.length; vi += 1) sub.push(valueLabel(vi));
    grid.push(sub);
    rowMeta.push({ kind: 'header' });
  }
  const headerRowCount = colFields.length + (multiValue ? 1 : 0);

  // ---- Value rows --------------------------------------------------
  // For each column tuple, slice the subset to records matching every column
  // field, aggregate each value; finish with the across-all grand-total block.
  const valueCellsFor = (records: PivotCell[][]): PivotCell[] => {
    const cells: PivotCell[] = [];
    for (const t of tuples) {
      const slice = records.filter((rec) => colFields.every((c, i) => keyOf(rec[c]) === t[i]));
      for (const v of values) {
        cells.push(
          aggregate(
            slice.map((r) => r[v.column]),
            v.agg,
          ),
        );
      }
    }
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

  const root = buildTree(filteredRecords, model.rows);
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

  // "Show Values As → % of Grand Total" for the cross-tab layout: every value
  // cell of a flagged field becomes its share of that field's overall grand
  // total (the bottom-right cell — last row's grand-total column for the field).
  applyMatrixShowAsPercent(grid, colMeta, values, headerRowCount);

  return { grid, rowMeta, colMeta };
}

/** Cross-tab analogue of {@link applyShowAsPercent}: divide every value cell of
 *  each `pctOfGrandTotal` field by that field's overall grand total (the bottom
 *  row's grand-total column). Mutates `grid` in place; header rows are skipped. */
function applyMatrixShowAsPercent(
  grid: PivotGrid,
  colMeta: PivotColMeta[],
  values: PivotValueField[],
  headerRows: number,
): void {
  const lastRow = grid.length - 1;
  if (lastRow < headerRows) return;
  values.forEach((v, vi) => {
    const mode = v.showAs;
    if (mode !== 'pctOfGrandTotal' && mode !== 'pctOfColumnTotal' && mode !== 'pctOfRowTotal') {
      return;
    }
    const cols = colMeta
      .map((m, i) => (m.kind !== 'label' && m.valueIndex === vi ? i : -1))
      .filter((i) => i >= 0);
    // Denominator per mode:
    //   Grand  → the field's bottom-right total (one value).
    //   Column → each column's bottom total (per c; bottom row is touched last).
    //   Row    → each row's grand-total cell (per r; read before mutating the row).
    const gtCol = colMeta.findIndex((m) => m.kind === 'grand-total' && m.valueIndex === vi);
    const grandDenom = gtCol >= 0 ? Number(grid[lastRow][gtCol]) : 0;
    for (let r = headerRows; r < grid.length; r++) {
      const rowDenom = mode === 'pctOfRowTotal' && gtCol >= 0 ? Number(grid[r][gtCol]) : 0;
      for (const c of cols) {
        const denom =
          mode === 'pctOfColumnTotal'
            ? Number(grid[lastRow][c])
            : mode === 'pctOfRowTotal'
              ? rowDenom
              : grandDenom;
        const n = Number(grid[r][c]);
        grid[r][c] = denom && Number.isFinite(n) ? `${((n / denom) * 100).toFixed(1)}%` : '0.0%';
      }
    }
  });
}

/** Coerce a cell value to its string bucket key (null/empty → ''). */
function keyOf(v: PivotCell): string {
  return v == null ? '' : String(v);
}

/** Excel serial number → {year, month(0-11)} (UTC), inverting the 1900
 *  leap-bug offset our importer applies (parse-impl `excelSerialFromDate`). */
function partsFromSerial(serial: number): { y: number; m: number } {
  const base = Date.UTC(1900, 0, 1);
  let days = Math.floor(serial) - 1;
  if (serial >= 60) days -= 1; // undo Excel's fictitious 1900-02-29
  const d = new Date(base + days * 86_400_000);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() };
}

/** Resolve a cell value to {year, month(0-11)}, handling both shapes dates take
 *  in this app: an Excel serial number (imported xlsx) and a date string like
 *  "2025/01/10" or "2025-01-10" (a DATE() formula / typed date). Returns null
 *  when the value isn't a recognisable date. */
function dateParts(value: PivotCell): { y: number; m: number } | null {
  if (typeof value === 'number' && Number.isFinite(value)) return partsFromSerial(value);
  const s = typeof value === 'string' ? value.trim() : '';
  if (s === '') return null;
  if (/^\d+(\.\d+)?$/.test(s)) return partsFromSerial(Number(s)); // numeric string → serial
  const ymd = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/.exec(s);
  if (ymd) return { y: Number(ymd[1]), m: Number(ymd[2]) - 1 };
  const t = Date.parse(s);
  if (!Number.isNaN(t)) {
    const d = new Date(t);
    return { y: d.getUTCFullYear(), m: d.getUTCMonth() };
  }
  return null;
}

/** Bucket a row-field value for date grouping. 'none' (or a value that isn't a
 *  recognisable date) returns the raw key; year/quarter/month derive a period
 *  key. Keys sort lexicographically into chronological order (e.g. "2025-01" <
 *  "2025-02", "2025-Q1" < "2025-Q2"). */
export function dateGroupKey(value: PivotCell, grouping: DateGrouping): string {
  if (grouping === 'none') return keyOf(value);
  const parts = dateParts(value);
  if (!parts) return keyOf(value);
  const { y, m } = parts;
  if (grouping === 'year') return String(y);
  if (grouping === 'quarter') return `${y}-Q${Math.floor(m / 3) + 1}`;
  return `${y}-${String(m + 1).padStart(2, '0')}`; // month
}

type TreeNode = {
  records: PivotCell[][];
  children: Map<string, TreeNode>;
};

function buildTree(records: PivotCell[][], rowFields: PivotFieldRef[]): TreeNode {
  const root: TreeNode = { records: [], children: new Map() };
  for (const rec of records) {
    let node = root;
    node.records.push(rec);
    for (const field of rowFields) {
      const key = dateGroupKey(rec[field.column], field.grouping ?? 'none');
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
