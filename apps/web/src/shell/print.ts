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

import type { FUniver } from '@univerjs/core/facade';

export type PrintOrientation = 'portrait' | 'landscape';
export type PrintMarginPreset = 'narrow' | 'normal' | 'wide';

export type PrintOptions = {
  orientation: PrintOrientation;
  margins: PrintMarginPreset;
  /** A1-notation range to restrict the print to. Null / undefined =
   *  whole used range. Excel calls this "Print area". Persists with
   *  the other print preferences. */
  printArea?: string | null;
};

export const DEFAULT_PRINT_OPTIONS: PrintOptions = {
  orientation: 'portrait',
  margins: 'normal',
  printArea: null,
};

const MARGIN_MM: Record<PrintMarginPreset, number> = {
  narrow: 6.35,
  normal: 18,
  wide: 25,
};

const STORAGE_KEY = 'casual-sheets/print-options';

/**
 * Soft cell-count cap before we refuse to render the print HTML.
 *
 * The renderer builds one big `<table>` and writes it into an iframe
 * `srcdoc`. Past ~100k cells:
 *   - The HTML string itself grows past several MB (each `<td>` is
 *     ~30 bytes minimum + the value text).
 *   - The browser materialises one DOM node per cell. At 200k+ cells
 *     the tab heap hits ~500 MB and Chrome OOMs the renderer (the
 *     6 MB-xlsx crash reported in issue #50).
 *
 * The hard cap stops the user mid-Print with a clear "set a Print
 * Area" recovery message rather than crashing the tab. 100k is the
 * smallest number that doesn't reject everyday workbooks; raise it
 * (or lower the FUDGE) when a user with a beefier device asks.
 */
export const PRINT_CELL_LIMIT = 100_000;

export type PrintResult =
  | { ok: true; cellCount: number }
  | { ok: false; reason: 'empty' }
  | { ok: false; reason: 'too-large'; cellCount: number; limit: number };

/**
 * Read previously-chosen print options from localStorage. Returns the
 * defaults if nothing is stored or the stored value looks malformed.
 */
export function loadPrintOptions(): PrintOptions {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return DEFAULT_PRINT_OPTIONS;
    const parsed = JSON.parse(raw) as Partial<PrintOptions>;
    const orientation: PrintOrientation =
      parsed.orientation === 'landscape' ? 'landscape' : 'portrait';
    const margins: PrintMarginPreset =
      parsed.margins === 'narrow' || parsed.margins === 'wide' ? parsed.margins : 'normal';
    const printArea =
      typeof parsed.printArea === 'string' && parsed.printArea.trim()
        ? parsed.printArea.trim()
        : null;
    return { orientation, margins, printArea };
  } catch {
    return DEFAULT_PRINT_OPTIONS;
  }
}

export function savePrintOptions(options: PrintOptions): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(options));
  } catch {
    /* private mode / storage full — silent, options just won't persist */
  }
}

/**
 * Print the active sheet via the browser's native print dialog.
 *
 * Univer OSS doesn't ship a print renderer (page setup / print preview are
 * Pro-only). We do the cheap-but-correct thing: render the active sheet's
 * used range as a plain HTML table styled for print, inject it into a hidden
 * iframe, and trigger window.print() in that iframe. The browser handles
 * paper size, margins, page headers/footers natively.
 *
 * Loss: formulas show their evaluated value (Excel's normal print behavior),
 * column widths are approximated, drawings/charts are not included.
 */
export function printActiveSheet(
  api: FUniver,
  options: PrintOptions = DEFAULT_PRINT_OPTIONS,
): PrintResult {
  const wb = api.getActiveWorkbook();
  if (!wb) return { ok: false, reason: 'empty' };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const fws: any = wb.getActiveSheet();
  if (!fws) return { ok: false, reason: 'empty' };
  const ws = fws.getSheet();

  // Either the user's explicit Print Area (Excel "Set print area")
  // or the auto-detected used range. The print-area string is in
  // A1 notation (e.g. "A1:D20"); a parse failure falls back to the
  // used range with a console warning rather than printing nothing.
  let startRow = 0;
  let startCol = 0;
  let lastRow = ws.getLastRowWithContent?.() ?? 0;
  let lastCol = ws.getLastColumnWithContent?.() ?? 0;
  if (options.printArea) {
    const parsed = parsePrintArea(options.printArea);
    if (parsed) {
      startRow = parsed.startRow;
      startCol = parsed.startCol;
      lastRow = parsed.endRow;
      lastCol = parsed.endCol;
    } else {
      console.warn(
        '[print] could not parse Print Area "%s" — falling back to used range',
        options.printArea,
      );
    }
  }
  if (lastRow < 0 || lastCol < 0) return { ok: false, reason: 'empty' };

  // Cell-count guard. The renderer builds one big HTML table; past
  // ~PRINT_CELL_LIMIT cells the tab OOMs (issue #50). Refuse rather
  // than crash; the caller surfaces a toast pointing at "Set a Print
  // Area first" as the recovery path. The limit can be overridden
  // via `window.__casualPrintCellLimit` — e2e specs use this to
  // exercise the soft-guard without standing up a multi-MB workbook.
  const rowCount = lastRow - startRow + 1;
  const colCount = lastCol - startCol + 1;
  const cellCount = rowCount * colCount;
  const overrideLimit =
    typeof window !== 'undefined'
      ? (window as Window & { __casualPrintCellLimit?: number }).__casualPrintCellLimit
      : undefined;
  const limit = typeof overrideLimit === 'number' ? overrideLimit : PRINT_CELL_LIMIT;
  if (cellCount > limit) {
    return {
      ok: false,
      reason: 'too-large',
      cellCount,
      limit,
    };
  }

  const sheetName: string = fws.getSheetName?.() ?? 'Sheet';
  const workbookName: string = (wb as { getName?: () => string }).getName?.() ?? 'Workbook';

  const rows: string[] = [];
  for (let r = startRow; r <= lastRow; r++) {
    const cells: string[] = [];
    for (let c = startCol; c <= lastCol; c++) {
      const cell = ws.getCell(r, c);
      const raw: unknown = cell?.v;
      const text =
        raw === undefined || raw === null ? '' : typeof raw === 'string' ? raw : String(raw);
      cells.push(`<td>${escapeHtml(text)}</td>`);
    }
    rows.push(`<tr>${cells.join('')}</tr>`);
  }

  // Render column-width hints in CSS so the print layout doesn't squash
  // narrow columns. ExcelJS-style units don't matter here — relative
  // proportions are what carry through to print.
  const colWidths: string[] = [];
  for (let c = startCol; c <= lastCol; c++) {
    const w = ws.getColumnWidth?.(c) ?? 88;
    colWidths.push(`<col style="width:${Math.round(w)}px" />`);
  }

  const marginMm = MARGIN_MM[options.margins];
  const pageRule = `@page { size: A4 ${options.orientation}; margin: ${marginMm}mm; }`;

  const html = `<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<title>${escapeHtml(workbookName)} — ${escapeHtml(sheetName)}</title>
<style>
  ${pageRule}
  body { font: 11pt Inter, system-ui, sans-serif; color: #1a1a1a; margin: 0; }
  h1 { font-size: 14pt; margin: 0 0 8pt; font-weight: 600; }
  .meta { color: #6b7280; font-size: 9pt; margin-bottom: 8pt; }
  table { border-collapse: collapse; width: 100%; }
  td { border: 1px solid #cbd5e1; padding: 4px 6px; vertical-align: top;
       font-size: 10pt; word-wrap: break-word; }
  thead td { background: #f1f5f9; font-weight: 600; }
  tr:nth-child(even) td { background: #f9fafb; }
</style>
</head>
<body>
  <h1>${escapeHtml(workbookName)} — ${escapeHtml(sheetName)}</h1>
  <div class="meta">Printed ${new Date().toLocaleString()}</div>
  <table>
    <colgroup>${colWidths.join('')}</colgroup>
    <tbody>${rows.join('')}</tbody>
  </table>
</body>
</html>`;

  // Hidden iframe avoids leaving the user's tab. Some browsers block
  // printing inside an iframe whose contentDocument has cross-origin
  // restrictions; writing via `srcdoc` keeps everything same-origin.
  const frame = document.createElement('iframe');
  frame.style.position = 'fixed';
  frame.style.right = '0';
  frame.style.bottom = '0';
  frame.style.width = '0';
  frame.style.height = '0';
  frame.style.border = '0';
  frame.srcdoc = html;
  document.body.appendChild(frame);

  frame.addEventListener(
    'load',
    () => {
      try {
        const win = frame.contentWindow;
        if (!win) return;
        win.focus();
        win.print();
      } finally {
        // Give the browser a moment to open its print dialog before we
        // detach the iframe; some engines abort the dialog otherwise.
        setTimeout(() => frame.remove(), 1000);
      }
    },
    { once: true },
  );
  return { ok: true, cellCount };
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** Parse a "Print Area" A1 range string. Accepts `A1:D20` or a single
 *  cell `A1`. Returns null for malformed input. */
export function parsePrintArea(s: string): {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
} | null {
  const trimmed = s.trim().toUpperCase();
  if (!trimmed) return null;
  const m = /^(\$?)([A-Z]+)(\$?)(\d+)(?::(\$?)([A-Z]+)(\$?)(\d+))?$/.exec(trimmed);
  if (!m) return null;
  const col1 = colLettersToIndex(m[2]);
  const row1 = parseInt(m[4], 10) - 1;
  const col2 = m[6] ? colLettersToIndex(m[6]) : col1;
  const row2 = m[8] ? parseInt(m[8], 10) - 1 : row1;
  if (Number.isNaN(row1) || Number.isNaN(row2)) return null;
  return {
    startRow: Math.min(row1, row2),
    endRow: Math.max(row1, row2),
    startCol: Math.min(col1, col2),
    endCol: Math.max(col1, col2),
  };
}

function colLettersToIndex(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i += 1) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n - 1;
}
