import ExcelJS from 'exceljs';
import {
  CustomRangeType,
  LocaleType,
  type ICellData,
  type IRange,
  type IStyleData,
  type IWorkbookData,
} from '@univerjs/core';
import { excelStyleToUniver } from './style-mapping';
import { INITIAL_COLUMNS, INITIAL_ROWS, UNIVER_VERSION } from '../snapshot';
import { RESOURCES_SHEET } from './constants';
import type { ImportedWorkbook } from './import';

/**
 * Pure conversion: ExcelJS workbook → Univer IWorkbookData snapshot.
 * Imported by both `parser.worker.ts` (where it actually runs) and the
 * type-only `import.ts` (for the public types). Splitting this out keeps
 * ExcelJS — which is large — out of the main bundle; only the worker
 * chunk pays for it.
 *
 * Behavior is byte-for-byte equivalent to the previous in-place impl;
 * see import.ts header for the fidelity scope.
 */

let hyperlinkIdCounter = 0;
const nextHyperlinkId = () =>
  `hl-${Date.now().toString(36)}-${(hyperlinkIdCounter++).toString(36)}`;

/**
 * Build a rich-text doc body that encodes a hyperlink at the cell. The
 * shape mirrors what `AddHyperLinkCommand` writes into `cell.p.body`
 * (see vendor: sheets-hyper-link's AddHyperLinkCommand impl). Putting
 * hyperlinks here at import time avoids the previous per-link serial
 * `executeCommand` round-trip after mount — for a workbook with
 * thousands of links that path took multiple seconds.
 */
function buildHyperlinkBody(display: string, url: string, id: string): ICellData['p'] {
  // Univer doc bodies end with `\r\n` (paragraph + section break). The
  // customRange covers only the visible text [0, display.length - 1].
  const dataStream = `${display}\r\n`;
  return {
    id: '__INTERNAL_EDITOR__DOCS_NORMAL',
    documentStyle: {},
    body: {
      dataStream,
      customRanges: [
        {
          startIndex: 0,
          endIndex: display.length - 1,
          rangeType: CustomRangeType.HYPERLINK,
          rangeId: id,
          properties: { url },
        },
      ],
      paragraphs: [{ startIndex: display.length }],
      sectionBreaks: [{ startIndex: display.length + 1 }],
      textRuns: [],
    },
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;
}

function readResourcesSheet(ws: ExcelJS.Worksheet): IWorkbookData['resources'] {
  const parts: string[] = [];
  ws.eachRow({ includeEmpty: false }, (row) => {
    const v = row.getCell(1).value;
    if (typeof v === 'string') parts.push(v);
  });
  if (parts.length === 0) return undefined;
  try {
    const parsed = JSON.parse(parts.join(''));
    if (Array.isArray(parsed)) return parsed as IWorkbookData['resources'];
  } catch {
    /* corrupt blob — drop silently, the workbook still opens */
  }
  return undefined;
}

function lettersToCol(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col - 1;
}

export async function workbookFromExcelJs(buffer: ArrayBuffer): Promise<ImportedWorkbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  const id = `wb-${Date.now()}`;
  const sheetOrder: string[] = [];
  const sheets: IWorkbookData['sheets'] = {};
  const styles: Record<string, IStyleData | null> = {};
  let styleCounter = 0;

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

  const PX_PER_CHAR = 7;
  const charsToPx = (chars: number) => Math.round(chars * PX_PER_CHAR);
  const pointsToPx = (pt: number) => Math.round((pt * 96) / 72);

  let resources: IWorkbookData['resources'] | undefined;
  for (const ws of wb.worksheets) {
    if (ws.name === RESOURCES_SHEET) {
      resources = readResourcesSheet(ws);
      break;
    }
  }

  for (const ws of wb.worksheets) {
    if (ws.name === RESOURCES_SHEET) continue;
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
        const r = rowNumber - 1;
        const c = colNumber - 1;
        maxRow = Math.max(maxRow, r);
        maxCol = Math.max(maxCol, c);

        const cd: ICellData = {};
        const raw = cell.value;
        if (raw && typeof raw === 'object' && 'formula' in raw) {
          const f = (raw as { formula: string }).formula;
          cd.f = f.startsWith('=') ? f : `=${f}`;
          const result = (raw as { result?: unknown }).result;
          if (result !== undefined && result !== null && typeof result !== 'object') {
            cd.v = result as ICellData['v'];
          }
        } else if (raw && typeof raw === 'object' && 'richText' in raw) {
          cd.v = (raw as { richText: { text: string }[] }).richText.map((t) => t.text).join('');
        } else if (raw && typeof raw === 'object' && 'text' in raw && 'hyperlink' in raw) {
          const display = (raw as { text: string }).text ?? '';
          cd.v = display;
          const url = (raw as { hyperlink: string }).hyperlink;
          if (typeof url === 'string' && url && display.length > 0) {
            // Inline the link into cell.p so it ships in the snapshot.
            // No need to replay through AddHyperLinkCommand at mount.
            cd.p = buildHyperlinkBody(display, url, nextHyperlinkId());
          }
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

    const merges = (ws.model as { merges?: string[] }).merges ?? [];
    for (const range of merges) {
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

    const wsColumns = (ws as { columns?: Array<{ width?: number; hidden?: boolean } | null> })
      .columns ?? [];
    wsColumns.forEach((col, i) => {
      if (!col) return;
      const entry: { w?: number; hd?: number } = {};
      if (typeof col.width === 'number') entry.w = charsToPx(col.width);
      if (col.hidden) entry.hd = 1;
      if (entry.w !== undefined || entry.hd !== undefined) columnData[i] = entry;
    });

    ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const entry: { h?: number; hd?: number } = {};
      if (typeof row.height === 'number') entry.h = pointsToPx(row.height);
      if (row.hidden) entry.hd = 1;
      if (entry.h !== undefined || entry.hd !== undefined) rowData[rowNumber - 1] = entry;
    });

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
          startRow: ySplit > 0 ? ySplit : -1,
          startColumn: xSplit > 0 ? xSplit : -1,
        };
      }
    }

    const argb = ws.properties?.tabColor?.argb;
    const tabColor =
      argb && /^[0-9A-Fa-f]{8}$/.test(argb) ? `#${argb.slice(2).toLowerCase()}` : undefined;

    const defaultColumnWidth =
      typeof ws.properties?.defaultColWidth === 'number'
        ? charsToPx(ws.properties.defaultColWidth)
        : undefined;
    const defaultRowHeight =
      typeof ws.properties?.defaultRowHeight === 'number'
        ? pointsToPx(ws.properties.defaultRowHeight)
        : undefined;

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
    ...(resources ? { resources } : {}),
  };
}
