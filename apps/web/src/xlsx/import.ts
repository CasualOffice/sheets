import ExcelJS from 'exceljs';
import { LocaleType, type ICellData, type IRange, type IStyleData, type IWorkbookData } from '@univerjs/core';
import { excelStyleToUniver } from './style-mapping';
import { INITIAL_COLUMNS, INITIAL_ROWS, UNIVER_VERSION } from '../snapshot';

/**
 * Convert an .xlsx buffer to a Univer `IWorkbookData` snapshot.
 *
 * Fidelity scope (MVP):
 *   - Values + formulas (cell.value / cell.formula)
 *   - Font (family, size, bold, italic, underline, color)
 *   - Fill (solid background)
 *   - Alignment (horizontal, vertical, wrap)
 *   - Number format
 *   - Borders (thin, per side, color preserved)
 *   - Merges
 *   - Sheet order + names
 *
 * Accepts loss: charts, drawings, pivots, validation, conditional formatting,
 * data tables, comments, hyperlinks, advanced borders (dashed/double), themes.
 */
export async function xlsxToWorkbookData(buffer: ArrayBuffer): Promise<IWorkbookData> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  const id = `wb-${Date.now()}`;
  const sheetOrder: string[] = [];
  const sheets: IWorkbookData['sheets'] = {};
  const styles: Record<string, IStyleData | null> = {};
  let styleCounter = 0;

  // Intern equivalent styles so each unique style maps to a single id.
  // ExcelJS cells with the same effective style still produce different JSON,
  // so we hash the canonical form.
  const styleByKey = new Map<string, string>();
  const internStyle = (style: IStyleData | undefined): string | undefined => {
    if (!style) return undefined;
    const key = JSON.stringify(style);
    const existing = styleByKey.get(key);
    if (existing) return existing;
    const styleId = `s${styleCounter++}`;
    styleByKey.set(key, styleId);
    styles[styleId] = style;
    return styleId;
  };

  // Same conversion constants as the exporter — keep them in lockstep so
  // a save → reopen → save cycle is a fixed point.
  const PX_PER_CHAR = 7;
  const charsToPx = (chars: number) => Math.round(chars * PX_PER_CHAR);
  const pointsToPx = (pt: number) => Math.round((pt * 96) / 72);

  for (const ws of wb.worksheets) {
    const sheetId = `sheet-${ws.id}`;
    sheetOrder.push(sheetId);

    const cellData: Record<number, Record<number, ICellData>> = {};
    const mergeData: IRange[] = [];
    const columnData: Record<number, { w?: number; hd?: number }> = {};
    const rowData: Record<number, { h?: number; hd?: number }> = {};

    let maxRow = 0;
    let maxCol = 0;

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        // ExcelJS is 1-indexed; Univer is 0-indexed.
        const r = rowNumber - 1;
        const c = colNumber - 1;
        maxRow = Math.max(maxRow, r);
        maxCol = Math.max(maxCol, c);

        const cd: ICellData = {};
        // value can be a primitive, a formula object {formula, result},
        // a rich text object, or a hyperlink object. Normalize:
        const raw = cell.value;
        if (raw && typeof raw === 'object' && 'formula' in raw) {
          // formula cell
          const f = (raw as { formula: string }).formula;
          cd.f = f.startsWith('=') ? f : `=${f}`;
          const result = (raw as { result?: unknown }).result;
          if (result !== undefined && result !== null && typeof result !== 'object') {
            cd.v = result as ICellData['v'];
          }
        } else if (raw && typeof raw === 'object' && 'richText' in raw) {
          cd.v = (raw as { richText: { text: string }[] }).richText.map((t) => t.text).join('');
        } else if (raw && typeof raw === 'object' && 'text' in raw && 'hyperlink' in raw) {
          cd.v = (raw as { text: string }).text;
        } else if (raw && typeof raw === 'object' && 'sharedFormula' in raw) {
          const sf = (raw as { sharedFormula: string; result?: unknown }).sharedFormula;
          cd.f = sf.startsWith('=') ? sf : `=${sf}`;
          const result = (raw as { result?: unknown }).result;
          if (result !== undefined && result !== null && typeof result !== 'object') {
            cd.v = result as ICellData['v'];
          }
        } else if (raw instanceof Date) {
          cd.v = raw.toISOString();
        } else if (typeof raw === 'number' || typeof raw === 'boolean' || typeof raw === 'string') {
          cd.v = raw;
        }

        const styleId = internStyle(excelStyleToUniver(cell));
        if (styleId) cd.s = styleId;

        if (cd.v !== undefined || cd.f || cd.s) {
          cellData[r] ??= {};
          cellData[r][c] = cd;
        }
      });
    });

    // Merges live as a record keyed by the top-left cell address e.g. "A1:B2".
    const merges = (ws.model as { merges?: string[] }).merges ?? [];
    for (const range of merges) {
      // "A1:B2"
      const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(range);
      if (!m) continue;
      const [, startColL, startRowS, endColL, endRowS] = m;
      const start = { row: Number(startRowS) - 1, col: lettersToCol(startColL) };
      const end = { row: Number(endRowS) - 1, col: lettersToCol(endColL) };
      mergeData.push({
        startRow: start.row,
        startColumn: start.col,
        endRow: end.row,
        endColumn: end.col,
      });
    }

    // Column widths (ExcelJS uses character units; convert to pixels).
    // ws.columns is undefined when no column metadata is set.
    const wsColumns = (ws as { columns?: Array<{ width?: number; hidden?: boolean } | null> })
      .columns ?? [];
    wsColumns.forEach((col, i) => {
      if (!col) return;
      const entry: { w?: number; hd?: number } = {};
      if (typeof col.width === 'number') entry.w = charsToPx(col.width);
      if (col.hidden) entry.hd = 1;
      if (entry.w !== undefined || entry.hd !== undefined) columnData[i] = entry;
    });

    // Row heights (ExcelJS uses points; convert to pixels). Only rows with
    // an explicit height are emitted by ExcelJS — empty rows are skipped.
    ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const entry: { h?: number; hd?: number } = {};
      if (typeof row.height === 'number') entry.h = pointsToPx(row.height);
      if (row.hidden) entry.hd = 1;
      if (entry.h !== undefined || entry.hd !== undefined) rowData[rowNumber - 1] = entry;
    });

    // Frozen panes — ExcelJS stores the first view; we read xSplit/ySplit.
    let freeze: { xSplit: number; ySplit: number; startRow: number; startColumn: number } | undefined;
    const view = (ws as { views?: Array<{ state?: string; xSplit?: number; ySplit?: number }> })
      .views?.[0];
    if (view?.state === 'frozen') {
      const xSplit = view.xSplit ?? 0;
      const ySplit = view.ySplit ?? 0;
      if (xSplit > 0 || ySplit > 0) {
        freeze = {
          xSplit,
          ySplit,
          // Univer's freeze.startRow/startColumn is the first non-frozen cell.
          startRow: ySplit > 0 ? ySplit : -1,
          startColumn: xSplit > 0 ? xSplit : -1,
        };
      }
    }

    // Tab color — ExcelJS gives { argb }, Univer wants '#rrggbb'.
    const argb = ws.properties?.tabColor?.argb;
    const tabColor =
      argb && /^[0-9A-Fa-f]{8}$/.test(argb) ? `#${argb.slice(2).toLowerCase()}` : undefined;

    // Sheet-level defaults (in our pixel units).
    const defaultColumnWidth =
      typeof ws.properties?.defaultColWidth === 'number'
        ? charsToPx(ws.properties.defaultColWidth)
        : undefined;
    const defaultRowHeight =
      typeof ws.properties?.defaultRowHeight === 'number'
        ? pointsToPx(ws.properties.defaultRowHeight)
        : undefined;

    // Hidden — ExcelJS exposes ws.state as 'visible' | 'hidden' | 'veryHidden'.
    const hidden = (ws as { state?: string }).state === 'hidden' ? 1 : undefined;

    sheets[sheetId] = {
      id: sheetId,
      name: ws.name,
      cellData,
      mergeData,
      columnData,
      rowData,
      rowCount: Math.max(INITIAL_ROWS, maxRow + 1),
      columnCount: Math.max(INITIAL_COLUMNS, maxCol + 1),
      ...(freeze ? { freeze } : {}),
      ...(tabColor ? { tabColor } : {}),
      ...(defaultColumnWidth !== undefined ? { defaultColumnWidth } : {}),
      ...(defaultRowHeight !== undefined ? { defaultRowHeight } : {}),
      ...(hidden ? { hidden } : {}),
    };
  }

  // If the file had no worksheets (rare but possible), seed with an empty one
  // so the loader doesn't blow up.
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
    name: wb.title || 'Untitled',
    appVersion: UNIVER_VERSION,
    locale: LocaleType.EN_US,
    styles,
    sheetOrder,
    sheets,
  };
}

function lettersToCol(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col - 1;
}
