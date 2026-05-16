import ExcelJS from 'exceljs';
import { LocaleType, type ICellData, type IRange, type IStyleData, type IWorkbookData } from '@univerjs/core';
import { excelStyleToUniver } from './style-mapping';
import { INITIAL_COLUMNS, INITIAL_ROWS } from '../snapshot';

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

  for (const ws of wb.worksheets) {
    const sheetId = `sheet-${ws.id}`;
    sheetOrder.push(sheetId);

    const cellData: Record<number, Record<number, ICellData>> = {};
    const mergeData: IRange[] = [];

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

    sheets[sheetId] = {
      id: sheetId,
      name: ws.name,
      cellData,
      mergeData,
      rowCount: Math.max(INITIAL_ROWS, maxRow + 1),
      columnCount: Math.max(INITIAL_COLUMNS, maxCol + 1),
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
    appVersion: '0.22.1',
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
