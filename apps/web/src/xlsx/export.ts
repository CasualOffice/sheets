import ExcelJS from 'exceljs';
import type { IStyleData, IWorkbookData } from '@univerjs/core';
import { univerStyleToExcel } from './style-mapping';

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
 * Convert a Univer `IWorkbookData` snapshot to an .xlsx Blob.
 * See `import.ts` for the fidelity scope (same coverage in both directions).
 */
export async function workbookDataToXlsx(data: IWorkbookData): Promise<Blob> {
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
  }

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}
