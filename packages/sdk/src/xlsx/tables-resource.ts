import type ExcelJS from 'exceljs';
import type { IWorkbookData } from '@univerjs/core';

/**
 * Passthrough preservation for xlsx ListObjects (a.k.a. Excel Tables).
 * Univer doesn't model tables as first-class objects — autofilters and
 * structured-reference parsing aren't there yet — so without a
 * passthrough, opening an xlsx that ships a defined table and saving
 * it again would silently strip the table back to plain cells.
 *
 * Data lives in our own sidecar (`__casual_sheets_tables__`); on
 * export we replay `ws.addTable(...)` so the saved file shows the
 * table to any reader (real Excel, gsheets, LibreOffice).
 *
 * When/if Univer grows native table support, this resource is the
 * place to switch from sidecar to plugin shape — readers of older
 * files keep working because the sidecar is forward-compatible.
 */

export const TABLES_RESOURCE = '__casual_sheets_tables__';

export type SynthTableColumn = {
  name: string;
  filterButton?: boolean;
  totalsRowFunction?: string;
  totalsRowLabel?: string;
  totalsRowFormula?: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  style?: any;
};

export type SynthTableV1 = {
  name: string;
  ref: string; // 'A1' or 'A1:C4'
  displayName?: string;
  headerRow?: boolean;
  totalsRow?: boolean;
  style?: {
    theme?: string;
    name?: string;
    showRowStripes?: boolean;
    showColumnStripes?: boolean;
    showFirstColumn?: boolean;
    showLastColumn?: boolean;
  };
  columns: SynthTableColumn[];
  rows: unknown[][];
};

type Payload = {
  v: 1;
  sheets: Record<string, SynthTableV1[]>;
};

/**
 * Walk every worksheet and lift any defined tables into the sidecar
 * payload. ExcelJS stores tables on `ws.tables` keyed by table name;
 * each value is a Table *instance* whose actual fields hang off
 * `t.model` (`node_modules/exceljs/lib/doc/table.js:275`). Two
 * gotchas the parser must handle:
 *
 *   1. After `xlsx.load`, the model uses `tableRef` (full range) and
 *      drops `ref` (the original top-left author hint). We coalesce
 *      onto `ref` as a single canonical full-range string.
 *   2. The loaded model never carries `rows` — table cell data is
 *      already on the worksheet. We don't need rows to round-trip.
 */
export function readTablesFromXlsx(
  wb: ExcelJS.Workbook,
  sheetIdForExcel: (excelId: number) => string,
): Record<string, SynthTableV1[]> {
  const out: Record<string, SynthTableV1[]> = {};
  for (const ws of wb.worksheets) {
    const tableMap = (ws as unknown as { tables?: Record<string, unknown> | Map<string, unknown> })
      .tables;
    if (!tableMap) continue;
    const entries: unknown[] = tableMap instanceof Map
      ? Array.from(tableMap.values())
      : Object.values(tableMap);
    if (entries.length === 0) continue;

    const tables: SynthTableV1[] = [];
    for (const tUnknown of entries) {
      const tInstance = tUnknown as { model?: unknown };
      const modelUnknown = (tInstance && typeof tInstance === 'object' && 'model' in tInstance)
        ? tInstance.model
        : tUnknown;
      const model = modelUnknown as Record<string, unknown>;
      if (!model || typeof model !== 'object') continue;
      const name = model.name;
      // Prefer the post-load `tableRef` (full range, eg `'A1:B3'`).
      // Fall back to `ref` for tables that were `addTable`d in-process
      // and never made the round-trip through xform load.
      const refStr = typeof model.tableRef === 'string'
        ? model.tableRef
        : typeof model.ref === 'string'
          ? model.ref
          : null;
      if (typeof name !== 'string' || !refStr) continue;

      const cols: SynthTableColumn[] = Array.isArray(model.columns)
        ? (model.columns as unknown[]).map((cUnknown) => {
            const c = cUnknown as Record<string, unknown>;
            return {
              name: typeof c?.name === 'string' ? c.name : '',
              ...(typeof c?.filterButton === 'boolean' ? { filterButton: c.filterButton } : {}),
              ...(typeof c?.totalsRowFunction === 'string'
                ? { totalsRowFunction: c.totalsRowFunction }
                : {}),
              ...(typeof c?.totalsRowLabel === 'string'
                ? { totalsRowLabel: c.totalsRowLabel }
                : {}),
              ...(typeof c?.totalsRowFormula === 'string'
                ? { totalsRowFormula: c.totalsRowFormula }
                : {}),
            };
          })
        : [];

      tables.push({
        name,
        ref: refStr,
        ...(typeof model.displayName === 'string' ? { displayName: model.displayName } : {}),
        ...(typeof model.headerRow === 'boolean' ? { headerRow: model.headerRow } : {}),
        ...(typeof model.totalsRow === 'boolean' ? { totalsRow: model.totalsRow } : {}),
        ...(model.style && typeof model.style === 'object'
          ? { style: model.style as SynthTableV1['style'] }
          : {}),
        columns: cols,
        rows: [],
      });
    }
    if (tables.length > 0) out[sheetIdForExcel(ws.id)] = tables;
  }
  return out;
}

/** Merge a synthesised table map into the snapshot resources. */
export function mergeTablesIntoResources(
  resources: IWorkbookData['resources'],
  payload: Record<string, SynthTableV1[]>,
): IWorkbookData['resources'] {
  if (Object.keys(payload).length === 0) return resources;
  const existing = resources?.find((r) => r.name === TABLES_RESOURCE);
  if (existing) return resources;
  const body: Payload = { v: 1, sheets: payload };
  const next = [...(resources ?? [])];
  next.push({ name: TABLES_RESOURCE, data: JSON.stringify(body) });
  return next;
}

/** Read the tables resource off a snapshot. */
export function readTablesFromSnapshot(
  data: IWorkbookData,
): Record<string, SynthTableV1[]> {
  const entry = data.resources?.find((r) => r.name === TABLES_RESOURCE);
  if (!entry?.data) return {};
  try {
    const parsed = JSON.parse(entry.data) as Partial<Payload>;
    if (parsed?.v !== 1 || !parsed.sheets) return {};
    return parsed.sheets;
  } catch {
    return {};
  }
}

/**
 * Apply a per-sheet table set onto an ExcelJS worksheet.
 *
 * Why this bypasses `ws.addTable(...)`:
 *
 *   `Table.validate()` (`node_modules/exceljs/lib/doc/table.js:129`)
 *   throws when `table.rows` is missing — but our round-trip case
 *   doesn't carry rows (cell data already lives in `cellData`). And
 *   `Table.store()` would *re-write* every cell from the rows
 *   array, clobbering whatever the user had set after the table was
 *   created. So instead of going through `addTable`, we attach the
 *   load-shape model directly the same way ExcelJS's worksheet xform
 *   setter does on load (`node_modules/exceljs/lib/doc/worksheet.js:
 *   917`). The serializer then walks `Object.values(ws.tables).map(t
 *   => t.model)` (`worksheet.js:856`) and writes the table parts —
 *   `target` + `id` are filled in by `xlsx.js:632` during prepare.
 */
export function applyTablesToXlsxWorksheet(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ws: any,
  tables: SynthTableV1[],
): void {
  if (!tables?.length) return;
  ws.tables = ws.tables ?? {};
  for (const t of tables) {
    if (!t?.name || !t?.ref || !Array.isArray(t.columns) || t.columns.length === 0) continue;
    const model: Record<string, unknown> = {
      name: t.name,
      displayName: t.displayName ?? t.name,
      tableRef: t.ref,
      autoFilterRef: t.ref,
      headerRow: t.headerRow ?? true,
      totalsRow: t.totalsRow ?? false,
      columns: t.columns.map((c) => ({
        name: c.name,
        ...(c.filterButton !== undefined ? { filterButton: c.filterButton } : {}),
        ...(c.totalsRowFunction ? { totalsRowFunction: c.totalsRowFunction } : {}),
        ...(c.totalsRowLabel ? { totalsRowLabel: c.totalsRowLabel } : {}),
        ...(c.totalsRowFormula ? { totalsRowFormula: c.totalsRowFormula } : {}),
      })),
      style: t.style ?? {},
    };
    ws.tables[t.name] = { model };
  }
}
