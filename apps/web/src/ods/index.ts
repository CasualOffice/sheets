/**
 * OpenDocument Spreadsheet (.ods) I/O.
 *
 * Univer's own xlsx codec (apps/web/src/xlsx/) stays the primary format —
 * this module exists so users can open files saved from LibreOffice Calc /
 * OpenOffice Calc / Google Sheets's .ods export and save back to the same
 * format. We use the SheetJS Community fork (`@e965/xlsx`, Apache-2.0)
 * because it's the only well-maintained library that handles .ods on npm.
 *
 * Scope (MVP):
 *   - Values + cached formula results
 *   - Sheet order + names
 *   - Merges
 *
 * Loss: styles, formulas (we keep cached values), column widths, charts,
 * frozen panes. SheetJS Community surfaces these in its model, but mapping
 * them through to Univer is a larger pass — comes in a follow-up.
 */
import * as XLSX from '@e965/xlsx';
import { LocaleType, type ICellData, type IRange, type IWorkbookData } from '@univerjs/core';
import { INITIAL_COLUMNS, INITIAL_ROWS, UNIVER_VERSION } from '../snapshot';

type TabularFormat = 'ods' | 'csv' | 'tsv';

function readWorkbook(buffer: ArrayBuffer, format: TabularFormat): XLSX.WorkBook {
  if (format === 'tsv') {
    const text = new TextDecoder().decode(new Uint8Array(buffer));
    return XLSX.read(text, { type: 'string', FS: '\t' });
  }
  if (format === 'csv') {
    const text = new TextDecoder().decode(new Uint8Array(buffer));
    return XLSX.read(text, { type: 'string' });
  }
  return XLSX.read(buffer, { type: 'array' });
}

export async function odsToWorkbookData(buffer: ArrayBuffer): Promise<IWorkbookData> {
  return tabularToWorkbookData(buffer, 'ods');
}

export async function csvToWorkbookData(buffer: ArrayBuffer): Promise<IWorkbookData> {
  return tabularToWorkbookData(buffer, 'csv');
}

export async function tsvToWorkbookData(buffer: ArrayBuffer): Promise<IWorkbookData> {
  return tabularToWorkbookData(buffer, 'tsv');
}

async function tabularToWorkbookData(
  buffer: ArrayBuffer,
  format: TabularFormat,
): Promise<IWorkbookData> {
  const wb = readWorkbook(buffer, format);
  const id = `wb-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  const sheetOrder: string[] = [];
  const sheets: IWorkbookData['sheets'] = {};

  for (const name of wb.SheetNames) {
    const sheet = wb.Sheets[name];
    if (!sheet || !sheet['!ref']) {
      const sheetId = `sheet-${sheetOrder.length + 1}`;
      sheetOrder.push(sheetId);
      sheets[sheetId] = {
        id: sheetId,
        name,
        cellData: {},
        rowCount: INITIAL_ROWS,
        columnCount: INITIAL_COLUMNS,
      };
      continue;
    }

    const sheetId = `sheet-${sheetOrder.length + 1}`;
    sheetOrder.push(sheetId);

    const range = XLSX.utils.decode_range(sheet['!ref']);
    const cellData: Record<number, Record<number, ICellData>> = {};
    let maxRow = 0;
    let maxCol = 0;

    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[addr];
        if (!cell) continue;

        const cd: ICellData = {};
        // SheetJS cell value lives in .v; the type code in .t (n/s/b/d/e).
        // For our level of fidelity we treat everything as the value with no
        // explicit type tag — Univer infers from the JS value.
        if (cell.v !== undefined && cell.v !== null) {
          if (cell.v instanceof Date) {
            cd.v = cell.v.toISOString();
          } else if (
            typeof cell.v === 'number' ||
            typeof cell.v === 'string' ||
            typeof cell.v === 'boolean'
          ) {
            cd.v = cell.v;
          } else {
            cd.v = String(cell.v);
          }
        }
        // Cached formulas come through as cell.f without the leading '='.
        if (typeof cell.f === 'string' && cell.f.length > 0) {
          cd.f = cell.f.startsWith('=') ? cell.f : `=${cell.f}`;
        }
        if (cd.v !== undefined || cd.f) {
          cellData[r] ??= {};
          cellData[r][c] = cd;
          if (r > maxRow) maxRow = r;
          if (c > maxCol) maxCol = c;
        }
      }
    }

    const mergeData: IRange[] = (sheet['!merges'] ?? []).map((m) => ({
      startRow: m.s.r,
      startColumn: m.s.c,
      endRow: m.e.r,
      endColumn: m.e.c,
    }));

    sheets[sheetId] = {
      id: sheetId,
      name,
      cellData,
      mergeData,
      rowCount: Math.max(INITIAL_ROWS, maxRow + 1),
      columnCount: Math.max(INITIAL_COLUMNS, maxCol + 1),
    };
  }

  if (sheetOrder.length === 0) {
    sheetOrder.push('sheet-1');
    sheets['sheet-1'] = {
      id: 'sheet-1',
      name: 'Sheet1',
      cellData: {},
      rowCount: INITIAL_ROWS,
      columnCount: INITIAL_COLUMNS,
    };
  }

  return {
    id,
    rev: 1,
    name: 'Untitled',
    appVersion: UNIVER_VERSION,
    locale: LocaleType.EN_US,
    styles: {},
    sheetOrder,
    sheets,
  };
}

export async function workbookDataToOds(data: IWorkbookData): Promise<Blob> {
  const wb = XLSX.utils.book_new();

  for (const sheetId of data.sheetOrder) {
    const wsd = data.sheets[sheetId];
    if (!wsd) continue;
    const sheet: XLSX.WorkSheet = {};

    const cellData = (wsd.cellData ?? {}) as Record<
      string,
      Record<string, { v?: string | number | boolean; f?: string }>
    >;
    let maxRow = 0;
    let maxCol = 0;
    for (const rKey of Object.keys(cellData)) {
      const r = Number(rKey);
      const row = cellData[rKey];
      for (const cKey of Object.keys(row)) {
        const c = Number(cKey);
        const cell = row[cKey];
        const addr = XLSX.utils.encode_cell({ r, c });
        const out: XLSX.CellObject = { t: 's', v: '' };
        if (cell.f) {
          out.f = cell.f.startsWith('=') ? cell.f.slice(1) : cell.f;
          if (cell.v !== undefined && cell.v !== null) {
            out.v = cell.v;
            out.t = typeof cell.v === 'number' ? 'n' : typeof cell.v === 'boolean' ? 'b' : 's';
          }
        } else if (cell.v !== undefined && cell.v !== null) {
          out.v = cell.v;
          out.t = typeof cell.v === 'number' ? 'n' : typeof cell.v === 'boolean' ? 'b' : 's';
        } else {
          continue;
        }
        sheet[addr] = out;
        if (r > maxRow) maxRow = r;
        if (c > maxCol) maxCol = c;
      }
    }

    if (Array.isArray(wsd.mergeData) && wsd.mergeData.length) {
      sheet['!merges'] = wsd.mergeData.map((m) => ({
        s: { r: m.startRow, c: m.startColumn },
        e: { r: m.endRow, c: m.endColumn },
      }));
    }

    sheet['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: maxRow, c: maxCol },
    });

    XLSX.utils.book_append_sheet(wb, sheet, wsd.name ?? sheetId);
  }

  const out = XLSX.write(wb, { type: 'array', bookType: 'ods' });
  return new Blob([out], {
    type: 'application/vnd.oasis.opendocument.spreadsheet',
  });
}

/**
 * Render the first sheet of the workbook as a CSV or TSV string. Both formats
 * are flat — they don't carry multi-sheet, styles, formulas (we emit cached
 * values), or merges. We export the active sheet only, matching what
 * LibreOffice / Excel do for CSV.
 */
export async function workbookDataToDelimited(
  data: IWorkbookData,
  format: 'csv' | 'tsv',
): Promise<Blob> {
  const firstId = data.sheetOrder[0];
  const wsd = data.sheets[firstId];
  if (!wsd) return new Blob([''], { type: 'text/plain' });

  const cellData = (wsd.cellData ?? {}) as Record<
    string,
    Record<string, { v?: string | number | boolean; f?: string }>
  >;
  let maxRow = 0;
  let maxCol = 0;
  for (const rKey of Object.keys(cellData)) {
    const r = Number(rKey);
    if (r > maxRow) maxRow = r;
    for (const cKey of Object.keys(cellData[rKey])) {
      const c = Number(cKey);
      if (c > maxCol) maxCol = c;
    }
  }

  const sep = format === 'tsv' ? '\t' : ',';
  const lines: string[] = [];
  for (let r = 0; r <= maxRow; r++) {
    const row = cellData[r];
    const cells: string[] = [];
    for (let c = 0; c <= maxCol; c++) {
      const cell = row?.[c];
      const raw = cell?.v;
      cells.push(escapeField(raw, sep));
    }
    lines.push(cells.join(sep));
  }

  const text = lines.join('\r\n');
  return new Blob([text], {
    type: format === 'csv' ? 'text/csv;charset=utf-8' : 'text/tab-separated-values;charset=utf-8',
  });
}

function escapeField(v: unknown, sep: string): string {
  if (v === undefined || v === null) return '';
  const s = typeof v === 'string' ? v : String(v);
  // CSV (and TSV with the same rules) require quoting when the field
  // contains the separator, a quote, or a line break.
  if (s.includes(sep) || s.includes('"') || s.includes('\n') || s.includes('\r')) {
    return `"${s.replace(/"/g, '""')}"`;
  }
  return s;
}
