/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * SheetsBridge — translates Sheets AI tool calls into FUniver facade operations.
 *
 * Read tools walk the active workbook/sheet and return JSON.
 * Write tools call FUniver range mutations (setValues / setValue).
 * The LLM never touches Univer internals directly.
 */

import type { FUniver } from '@univerjs/core/facade';
import { activeSheet, activeRange, rangeFromA1, rangeAt } from '../univer-facade';
import { retrieve, type RetrievalChunk } from './retrieval';

export type SheetsResult =
  | { ok: true; data?: unknown; diffSummary?: string }
  | { ok: false; code: string; message: string; retryable: boolean };

// ── A1 helpers ─────────────────────────────────────────────────────────────

function rowColToA1(row: number, col: number): string {
  let colStr = '';
  let c = col;
  do {
    colStr = String.fromCharCode((c % 26) + 65) + colStr;
    c = Math.floor(c / 26) - 1;
  } while (c >= 0);
  return `${colStr}${row + 1}`;
}

// ── Bridge ─────────────────────────────────────────────────────────────────

export class SheetsBridge {
  constructor(private readonly getApi: () => FUniver | null) {}

  async callTool(name: string, args: Record<string, unknown>): Promise<SheetsResult> {
    switch (name) {
      case 'get_workbook_info':
        return this.getWorkbookInfo();
      case 'get_selection':
        return this.getSelection();
      case 'get_cell_range':
        return this.getCellRange(args);
      case 'get_sheet_stats':
        return this.getSheetStats();
      case 'find_in_sheet':
        return this.findInSheet(args);
      case 'search_sheet':
        return this.searchSheet(args);
      case 'set_cell_values':
        return this.setCellValues(args);
      case 'set_formula':
        return this.setFormula(args);
      default:
        return {
          ok: false,
          code: 'UNSUPPORTED',
          message: `Unknown tool: ${name}`,
          retryable: false,
        };
    }
  }

  private noApi(): SheetsResult {
    return {
      ok: false,
      code: 'LOCATOR_NOT_FOUND',
      message: 'No active spreadsheet.',
      retryable: true,
    };
  }

  private getWorkbookInfo(): SheetsResult {
    const api = this.getApi();
    if (!api) return this.noApi();
    const wb = api.getActiveWorkbook();
    if (!wb) return this.noApi();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const activeWs = wb.getActiveSheet() as any;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const sheets = (wb.getSheets() as any[]).map((s: any) => ({
      id: s.getSheetId?.() ?? null,
      name: s.getSheetName?.() ?? s.getName?.() ?? null,
      rowCount: s.getMaxRows?.() ?? null,
      columnCount: s.getMaxColumns?.() ?? null,
      isActive: s === activeWs,
    }));

    return { ok: true, data: { sheets, sheetCount: sheets.length } };
  }

  private getSelection(): SheetsResult {
    const api = this.getApi();
    if (!api) return this.noApi();

    const range = activeRange(api);
    if (!range) return { ok: true, data: { hasSelection: false } };

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const r = range as any;
    const a1 = r.getA1Notation?.() ?? null;
    const values = r.getValues?.() ?? null;

    return { ok: true, data: { hasSelection: true, a1, values } };
  }

  private getCellRange(args: Record<string, unknown>): SheetsResult {
    const api = this.getApi();
    if (!api) return this.noApi();

    const rangeA1 = String(args.range ?? '');
    if (!rangeA1) {
      return { ok: false, code: 'VALIDATION', message: 'range is required.', retryable: false };
    }

    const sheet = activeSheet(api);
    if (!sheet) return this.noApi();

    const range = rangeFromA1(sheet, rangeA1);
    if (!range) {
      return {
        ok: false,
        code: 'VALIDATION',
        message: `Range "${rangeA1}" could not be resolved on the active sheet.`,
        retryable: false,
      };
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const r = range as any;
    const values = r.getValues?.() ?? null;
    const a1 = r.getA1Notation?.() ?? rangeA1;

    return { ok: true, data: { range: a1, values } };
  }

  private getSheetStats(): SheetsResult {
    const api = this.getApi();
    if (!api) return this.noApi();

    const sheet = activeSheet(api);
    if (!sheet) return this.noApi();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dataRange = (sheet as any).getDataRange?.();
    if (!dataRange) {
      return {
        ok: true,
        data: { rowCount: 0, columnCount: 0, nonEmptyCells: 0 },
      };
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const values = (dataRange as any).getValues?.() ?? [];
    const rowCount = values.length;
    const columnCount = (values[0] as unknown[])?.length ?? 0;
    let nonEmptyCells = 0;

    for (const row of values as unknown[][]) {
      for (const cell of row) {
        if (cell !== null && cell !== undefined && cell !== '') nonEmptyCells++;
      }
    }

    return { ok: true, data: { rowCount, columnCount, nonEmptyCells } };
  }

  private findInSheet(args: Record<string, unknown>): SheetsResult {
    const api = this.getApi();
    if (!api) return this.noApi();

    const query = String(args.query ?? '').toLowerCase();
    if (!query) {
      return { ok: false, code: 'VALIDATION', message: 'query is required.', retryable: false };
    }
    const limit = typeof args.limit === 'number' ? Math.min(args.limit, 50) : 20;

    const sheet = activeSheet(api);
    if (!sheet) return this.noApi();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const s = sheet as any;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dataRange = s.getDataRange?.() as any;
    if (!dataRange) return { ok: true, data: { matches: [], count: 0 } };

    const startRow: number = dataRange.getRow?.() ?? 0;
    const startCol: number = dataRange.getColumn?.() ?? 0;
    const values: unknown[][] = dataRange.getValues?.() ?? [];

    const matches: Array<{ cell: string; value: string }> = [];

    for (let r = 0; r < values.length && matches.length < limit; r++) {
      const row = values[r] ?? [];
      for (let c = 0; c < row.length && matches.length < limit; c++) {
        const cell = row[c];
        if (cell === null || cell === undefined || cell === '') continue;
        const cellStr = String(cell);
        if (!cellStr.toLowerCase().includes(query)) continue;

        const absRow = startRow + r;
        const absCol = startCol + c;
        // Try the facade first; fall back to manual A1 conversion.
        const cellRange = rangeAt(sheet, absRow, absCol);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const cellA1 = (cellRange as any)?.getA1Notation?.() ?? rowColToA1(absRow, absCol);
        matches.push({ cell: cellA1, value: cellStr.slice(0, 100) });
      }
    }

    return { ok: true, data: { matches, count: matches.length } };
  }

  /**
   * RAG: chunk the active sheet into header-carrying row bands and return the
   * bands most relevant to `query` (BM25), each with its A1 range so the model
   * can read the full band with get_cell_range before editing. Avoids sending
   * the whole (potentially huge) data range to the model.
   */
  private searchSheet(args: Record<string, unknown>): SheetsResult {
    const api = this.getApi();
    if (!api) return this.noApi();
    const query = String(args.query ?? '').trim();
    if (!query) {
      return { ok: false, code: 'VALIDATION', message: 'query is required.', retryable: false };
    }
    const k = typeof args.k === 'number' ? Math.min(Math.max(Math.floor(args.k), 1), 8) : 5;

    const sheet = activeSheet(api);
    if (!sheet) return this.noApi();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const s = sheet as any;
    const sheetName: string = s.getSheetName?.() ?? s.getName?.() ?? 'Sheet1';
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dataRange = s.getDataRange?.() as any;
    const values: unknown[][] = dataRange?.getValues?.() ?? [];
    if (!values.length) {
      return { ok: true, data: { chunks: [], count: 0, note: 'The sheet is empty.' } };
    }

    const startRow: number = dataRange.getRow?.() ?? 0;
    const startCol: number = dataRange.getColumn?.() ?? 0;
    const colLetter = (col: number) => rowColToA1(0, col).replace(/\d+$/, '');
    const rowToTsv = (row: unknown[]) => row.map((c) => c ?? '').join('\t');

    const header = values[0] ?? [];
    const headerTsv = rowToTsv(header);
    const lastCol = startCol + Math.max(header.length, 1) - 1;
    const BAND = 30;

    const chunks: RetrievalChunk[] = [];
    for (let start = 1; start < values.length; start += BAND) {
      const band = values.slice(start, start + BAND);
      const firstRow = startRow + start;
      const lastRow = startRow + Math.min(start + BAND, values.length) - 1;
      const a1 = `${sheetName}!${colLetter(startCol)}${firstRow + 1}:${colLetter(lastCol)}${lastRow + 1}`;
      chunks.push({
        id: `sc${chunks.length}`,
        // Header row carried in every band so column meaning is preserved.
        text: headerTsv + '\n' + band.map(rowToTsv).join('\n'),
        meta: { a1Range: a1 },
      });
    }

    const result = retrieve(chunks, query, { k });
    return {
      ok: true,
      data: {
        chunks: result.chunks.map((c) => ({
          a1Range: (c.meta as { a1Range?: string })?.a1Range ?? null,
          snippet: c.text.slice(0, 800),
          score: Math.round(c.score * 100) / 100,
        })),
        count: result.chunks.length,
        truncated: result.truncated,
        note: result.chunks.length
          ? 'Read a1Range with get_cell_range for full detail before editing.'
          : 'No rows matched the query.',
      },
    };
  }

  private setCellValues(args: Record<string, unknown>): SheetsResult {
    const api = this.getApi();
    if (!api) return this.noApi();

    const rangeA1 = String(args.range ?? '');
    const values = args.values;

    if (!rangeA1) {
      return { ok: false, code: 'VALIDATION', message: 'range is required.', retryable: false };
    }
    if (!Array.isArray(values)) {
      return {
        ok: false,
        code: 'VALIDATION',
        message: 'values must be a 2D array (array of arrays).',
        retryable: false,
      };
    }

    const sheet = activeSheet(api);
    if (!sheet) return this.noApi();

    const range = rangeFromA1(sheet, rangeA1);
    if (!range) {
      return {
        ok: false,
        code: 'VALIDATION',
        message: `Range "${rangeA1}" could not be resolved on the active sheet.`,
        retryable: false,
      };
    }

    // Coerce clean numeric-literal strings to numbers so "100" isn't stored as
    // text (which silently breaks SUM and other aggregations). Leave anything
    // that would change meaning alone: leading-zero codes ("007"), thousands
    // separators ("1,000"), currency ("$5"), etc.
    const coerced = (values as unknown[][]).map((row) =>
      Array.isArray(row)
        ? row.map((cell) => {
            if (typeof cell === 'string') {
              const t = cell.trim();
              if (/^-?\d+(\.\d+)?$/.test(t) && !/^-?0\d/.test(t)) return Number(t);
            }
            return cell;
          })
        : row,
    );

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const setValues = (range as any).setValues;
    if (typeof setValues !== 'function') {
      // Assert the facade method exists — optional chaining (setValues?.()) would
      // silently no-op and then report ok:true, so the model claims the write
      // happened when nothing was written.
      return {
        ok: false,
        code: 'INTERNAL',
        message: 'Spreadsheet write API unavailable (setValues missing).',
        retryable: false,
      };
    }
    try {
      setValues.call(range, coerced);
    } catch {
      return {
        ok: false,
        code: 'VALIDATION',
        message: 'Could not set values — range and values shape may not match.',
        retryable: false,
      };
    }

    return { ok: true, diffSummary: `Set values in ${rangeA1}.` };
  }

  private setFormula(args: Record<string, unknown>): SheetsResult {
    const api = this.getApi();
    if (!api) return this.noApi();

    const cellA1 = String(args.cell ?? '');
    const formula = String(args.formula ?? '');

    if (!cellA1) {
      return { ok: false, code: 'VALIDATION', message: 'cell is required.', retryable: false };
    }
    if (!formula.trim()) {
      return { ok: false, code: 'VALIDATION', message: 'formula is required.', retryable: false };
    }

    const sheet = activeSheet(api);
    if (!sheet) return this.noApi();

    const cell = rangeFromA1(sheet, cellA1);
    if (!cell) {
      return {
        ok: false,
        code: 'VALIDATION',
        message: `Cell "${cellA1}" could not be resolved on the active sheet.`,
        retryable: false,
      };
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const setValue = (cell as any).setValue;
    if (typeof setValue !== 'function') {
      return {
        ok: false,
        code: 'INTERNAL',
        message: 'Spreadsheet write API unavailable (setValue missing).',
        retryable: false,
      };
    }
    try {
      const f = formula.startsWith('=') ? formula : `=${formula}`;
      setValue.call(cell, { f });
    } catch {
      return {
        ok: false,
        code: 'VALIDATION',
        message: 'Could not set formula.',
        retryable: false,
      };
    }

    return { ok: true, diffSummary: `Set formula in ${cellA1}: ${formula}` };
  }
}
