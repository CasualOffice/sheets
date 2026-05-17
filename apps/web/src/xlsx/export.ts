import ExcelJS from 'exceljs';
import { timeItAsync } from '../perf';
import type { IStyleData, IWorkbookData } from '@univerjs/core';
import { univerStyleToExcel } from './style-mapping';
import { writeOutlineIntoSnapshot } from '../outline/resources';
import type { OutlineState } from '../outline/types';

type ICellSnapshot = {
  v?: string | number | boolean;
  f?: string;
  s?: string | IStyleData;
};

// Univer stores column width in pixels, Excel uses character widths
// based on the default workbook font. The Excel docs define width as
// "the number of characters of the largest digit (0-9) in the normal
// style's font that fit in the column," which empirically resolves to
// roughly 7 px per character at the default 11pt Calibri. ExcelJS exposes
// the same number, so we convert at the boundary.
const PX_PER_CHAR = 7;
const pxToChars = (px: number) => Math.max(0, px / PX_PER_CHAR);
// Univer stores row height in pixels, Excel uses points. 96dpi → 72pt.
const pxToPoints = (px: number) => Math.max(0, (px * 72) / 96);

/**
 * Cell-level extras the caller has read out of plugin services and wants us
 * to fold into the xlsx output. None of these live on `IWorkbookData` itself,
 * so the export function can't recover them on its own.
 */
export type ExportExtras = {
  /** subUnitId -> rows of { row, column, payload, display } */
  hyperlinks?: Record<string, Array<{ row: number; column: number; payload: string; display?: string }>>;
  /** Per-sheet row/column outline groups — survives the round-trip via two
   *  parallel channels: our `__casual_sheets_outline__` resource (exact
   *  group boundaries) AND ExcelJS row/col `outlineLevel`+`collapsed` (so
   *  Excel renders the native +/- gutter when the file is opened there). */
  outline?: OutlineState;
};

/**
 * Convert a Univer `IWorkbookData` snapshot to an .xlsx Blob.
 * See `import.ts` for the fidelity scope (same coverage in both directions).
 */
export async function workbookDataToXlsx(
  data: IWorkbookData,
  extras: ExportExtras = {},
): Promise<Blob> {
  return timeItAsync('export-xlsx', () => _workbookDataToXlsxImpl(data, extras));
}

async function _workbookDataToXlsxImpl(
  data: IWorkbookData,
  extras: ExportExtras = {},
): Promise<Blob> {
  const wb = new ExcelJS.Workbook();
  wb.title = data.name || 'Untitled';

  // Resolve a style ref (string id or inline IStyleData) to the IStyleData object.
  const resolveStyle = (s: ICellSnapshot['s']): IStyleData | undefined => {
    if (!s) return undefined;
    if (typeof s === 'string') return (data.styles?.[s] ?? undefined) as IStyleData | undefined;
    return s as IStyleData;
  };

  for (const sheetId of data.sheetOrder) {
    const wsd = data.sheets[sheetId];
    if (!wsd) continue;
    const ws = wb.addWorksheet(wsd.name ?? sheetId);

    // Hidden sheet — BooleanNumber 1 means hidden in Univer.
    if (wsd.hidden === 1) ws.state = 'hidden';

    // Tab color — Univer stores '#rrggbb' or 'rgb(...)'; ExcelJS wants ARGB hex.
    if (wsd.tabColor && typeof wsd.tabColor === 'string' && wsd.tabColor.startsWith('#')) {
      const rgb = wsd.tabColor.slice(1).toUpperCase();
      if (/^[0-9A-F]{6}$/.test(rgb)) {
        ws.properties.tabColor = { argb: `FF${rgb}` };
      }
    }

    // Frozen panes — Univer.freeze.ySplit/xSplit count the frozen rows/cols.
    if (wsd.freeze && (wsd.freeze.xSplit > 0 || wsd.freeze.ySplit > 0)) {
      ws.views = [{
        state: 'frozen',
        xSplit: wsd.freeze.xSplit || 0,
        ySplit: wsd.freeze.ySplit || 0,
      }];
    }

    // Sheet-level defaults.
    if (typeof wsd.defaultColumnWidth === 'number' && wsd.defaultColumnWidth > 0) {
      ws.properties.defaultColWidth = pxToChars(wsd.defaultColumnWidth);
    }
    if (typeof wsd.defaultRowHeight === 'number' && wsd.defaultRowHeight > 0) {
      ws.properties.defaultRowHeight = pxToPoints(wsd.defaultRowHeight);
    }

    const cellData = (wsd.cellData ?? {}) as Record<string, Record<string, ICellSnapshot>>;
    for (const rKey of Object.keys(cellData)) {
      const r = Number(rKey);
      const row = cellData[rKey];
      for (const cKey of Object.keys(row)) {
        const c = Number(cKey);
        const cell = row[cKey];

        // ExcelJS uses 1-indexed positions.
        const excelCell = ws.getCell(r + 1, c + 1);

        if (cell.f) {
          // Formula cell — strip leading '=' (ExcelJS adds it back).
          const formula = cell.f.startsWith('=') ? cell.f.slice(1) : cell.f;
          excelCell.value = { formula, result: cell.v ?? null } as ExcelJS.CellValue;
        } else if (cell.v !== undefined && cell.v !== null) {
          excelCell.value = cell.v as ExcelJS.CellValue;
        }

        const styleObj = resolveStyle(cell.s);
        if (styleObj) {
          Object.assign(excelCell, univerStyleToExcel(styleObj));
        }
      }
    }

    // Merges
    if (Array.isArray(wsd.mergeData)) {
      for (const m of wsd.mergeData) {
        ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1);
      }
    }

    // Hyperlinks for this sheet — write cell.value as { text, hyperlink }
    // so Excel and LibreOffice render them as native links. Applied AFTER
    // cellData so cell text we already set isn't clobbered.
    const sheetHyperlinks = extras.hyperlinks?.[sheetId] ?? [];
    for (const hl of sheetHyperlinks) {
      const excelCell = ws.getCell(hl.row + 1, hl.column + 1);
      const text =
        hl.display ??
        (typeof excelCell.value === 'string'
          ? excelCell.value
          : typeof excelCell.value === 'number'
            ? String(excelCell.value)
            : hl.payload);
      excelCell.value = { text, hyperlink: hl.payload };
    }

    // Column widths.
    const columnData = (wsd.columnData ?? {}) as Record<string, { w?: number; hd?: number }>;
    for (const cKey of Object.keys(columnData)) {
      const c = Number(cKey);
      const meta = columnData[cKey];
      if (typeof meta?.w === 'number' && meta.w > 0) {
        ws.getColumn(c + 1).width = pxToChars(meta.w);
      }
      if (meta?.hd === 1) ws.getColumn(c + 1).hidden = true;
    }

    // Row heights.
    const rowData = (wsd.rowData ?? {}) as Record<string, { h?: number; hd?: number }>;
    for (const rKey of Object.keys(rowData)) {
      const r = Number(rKey);
      const meta = rowData[rKey];
      if (typeof meta?.h === 'number' && meta.h > 0) {
        ws.getRow(r + 1).height = pxToPoints(meta.h);
      }
      if (meta?.hd === 1) ws.getRow(r + 1).hidden = true;
    }

    // Outline groups for this sheet — set row.outlineLevel / col.outlineLevel
    // (and collapsed where applicable) so Excel and LibreOffice render the
    // native +/- outline gutter. The full group state still round-trips
    // through our resource (extras.outline → data.resources below) so we
    // can reconstruct exact group ids and boundaries on re-import; the
    // outlineLevel is the "looks right in Excel" half of the bargain.
    const sheetOutline = extras.outline?.[sheetId];
    if (sheetOutline) {
      for (const g of sheetOutline.rows ?? []) {
        for (let r = g.start; r <= g.end; r++) {
          const row = ws.getRow(r + 1);
          row.outlineLevel = 1;
          if (g.collapsed) row.hidden = true;
        }
      }
      for (const g of sheetOutline.cols ?? []) {
        for (let c = g.start; c <= g.end; c++) {
          const col = ws.getColumn(c + 1);
          col.outlineLevel = 1;
          if (g.collapsed) col.hidden = true;
        }
      }
    }
  }

  // Fold outline state into the snapshot's resources before stashing them
  // into the hidden sheet — this is the channel that preserves exact group
  // ids + boundaries (vs `outlineLevel` which is just per-row/col flags).
  if (extras.outline && Object.keys(extras.outline).length > 0) {
    writeOutlineIntoSnapshot(data, extras.outline);
  }

  // Univer plugin state — tables, conditional formatting, data validation,
  // comments, notes, drawings, defined names, filters, range protection —
  // lives on the snapshot's `resources` array as serialized JSON per plugin.
  // None of it has a clean native xlsx representation that round-trips back
  // through our codec, so we stash the array in a hidden sheet with a known
  // name. On import we look for that sheet, decode, restore to data.resources,
  // and drop the sheet from sheetOrder before handing the snapshot to Univer.
  if (Array.isArray(data.resources) && data.resources.length > 0) {
    const meta = wb.addWorksheet(RESOURCES_SHEET);
    meta.state = 'veryHidden';
    // Split into chunks so Excel's 32k character-per-cell cap doesn't trip
    // for big workbooks.
    const json = JSON.stringify(data.resources);
    const CHUNK = 30_000;
    for (let i = 0, row = 1; i < json.length; i += CHUNK, row++) {
      meta.getCell(row, 1).value = json.slice(i, i + CHUNK);
    }
  }

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

/**
 * Name of the hidden sheet we stash Univer plugin resources in. Exported so
 * the importer can recognize and unpack it.
 */
export const RESOURCES_SHEET = '__casual_sheets_resources__';
