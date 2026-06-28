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

import { expect, test, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

type ParsedCell = {
  row: number;
  column: number;
  v?: unknown;
  s?: { bl?: number } | string;
  f?: unknown;
  rowSpan?: number;
  colSpan?: number;
};

type ParsedClipboardDump = {
  cells: ParsedCell[];
  rowProperties: Array<Record<string, unknown>>;
  colProperties: Array<Record<string, unknown>>;
};

/**
 * Paste-from-external-Excel fidelity. The internal copy/paste suite
 * already covers command-bus round-trips; this file targets the
 * Excel-flavored `text/html` parser and our post-parse merge hook.
 *
 * Headless Chromium's full DOM paste path is still flaky because
 * selection/render timing can make the visible grid assertions act
 * like a no-op even when the parser succeeds. So these tests hit the
 * clipboard service's HTML->USM conversion directly through a dev-only
 * helper and assert the exact cell/style/span/formula data that the
 * visible paste pipeline consumes.
 */

const EXCEL_HTML_BASIC = `
<meta http-equiv="content-type" content="text/html; charset=utf-8">
<table border="0" cellpadding="0" cellspacing="0">
  <tr height="20">
    <td style="font-weight:bold;color:#1f497d">Code</td>
    <td style="font-weight:bold;color:#1f497d">Qty</td>
  </tr>
  <tr height="20">
    <td>AAA</td>
    <td>10</td>
  </tr>
  <tr height="20">
    <td>BBB</td>
    <td>20</td>
  </tr>
</table>
`;

const EXCEL_HTML_MERGED = `
<table border="1">
  <tr>
    <td colspan="2">merged-header</td>
  </tr>
  <tr>
    <td>x</td>
    <td>y</td>
  </tr>
</table>
`;

const EXCEL_HTML_FORMULA = `
<table>
  <tr>
    <td>1</td>
    <td>2</td>
    <td x:fmla="=SUM(A1:B1)">3</td>
  </tr>
</table>
`;

async function parseHtml(page: Page, html: string) {
  const parsed = await page.evaluate((html) => {
    return window.__parseHtmlClipboard__?.(html) ?? null;
  }, html);
  expect(parsed).not.toBeNull();
  return parsed as ParsedClipboardDump;
}

function getCell(parsed: ParsedClipboardDump, row: number, column: number) {
  return parsed.cells.find((cell) => cell.row === row && cell.column === column) ?? null;
}

async function mergeMutationsFromParsed(page: Page, parsed: ParsedClipboardDump) {
  return page.evaluate(({ parsed }) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const hook = (window as any).__pasteMergeHook__ as
      | ((
          from: unknown,
          to: { range: { rows: number[]; cols: number[] }; unitId: string; subUnitId: string },
          matrix: {
            forValue: (
              cb: (
                row: number,
                column: number,
                cell: { rowSpan?: number; colSpan?: number } | null,
              ) => void,
            ) => void;
          },
          payload?: unknown,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
        ) => { undos: any[]; redos: any[] })
      | undefined;
    if (!hook) return null;
    const api = window.__univerAPI!;
    const wb = api.getActiveWorkbook()!;
    const pasteTo = {
      range: { rows: [0, 1, 2, 3], cols: [0, 1, 2, 3] },
      unitId: wb.getId(),
      subUnitId: wb.getActiveSheet()!.getSheetId(),
    };
    const matrix = {
      forValue(
        cb: (
          row: number,
          column: number,
          cell: { rowSpan?: number; colSpan?: number } | null,
        ) => void,
      ) {
        for (const cell of parsed.cells) {
          cb(cell.row, cell.column, {
            rowSpan: cell.rowSpan,
            colSpan: cell.colSpan,
          });
        }
      },
    };
    return hook(null, pasteTo, matrix, {});
  }, { parsed });
}

test.describe('Paste from external Excel (text/html clipboard)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('basic 2-column / 3-row table values land in the right cells', async ({ page }) => {
    const parsed = await parseHtml(page, EXCEL_HTML_BASIC);

    expect(String(getCell(parsed, 0, 0)?.v ?? '')).toBe('Code');
    expect(String(getCell(parsed, 0, 1)?.v ?? '')).toBe('Qty');
    expect(String(getCell(parsed, 2, 0)?.v ?? '')).toBe('BBB');
    expect(Number(getCell(parsed, 2, 1)?.v)).toBe(20);
  });

  test('header row keeps its bold style', async ({ page }) => {
    const parsed = await parseHtml(page, EXCEL_HTML_BASIC);
    const a1 = getCell(parsed, 0, 0);

    expect(typeof a1?.s).toBe('object');
    expect((a1?.s as { bl?: number } | undefined)?.bl).toBe(1);
  });

  test('merged-cell HTML preserves span metadata and emits a merge mutation', async ({ page }) => {
    const parsed = await parseHtml(page, EXCEL_HTML_MERGED);
    const a1 = getCell(parsed, 0, 0);

    expect(String(a1?.v ?? '')).toBe('merged-header');
    expect(a1?.colSpan).toBe(2);
    expect(a1?.rowSpan).toBe(1);

    const mutations = await mergeMutationsFromParsed(page, parsed);
    expect(mutations?.redos).toHaveLength(1);
    expect(mutations?.redos[0]?.id).toBe('sheet.mutation.add-worksheet-merge');
    expect(mutations?.redos[0]?.params?.ranges).toHaveLength(1);
    expect(mutations?.redos[0]?.params?.ranges[0]).toMatchObject({
      startRow: 0,
      endRow: 0,
      startColumn: 0,
      endColumn: 1,
    });
  });

  test('Excel formula in x:fmla attribute is preserved', async ({ page }) => {
    const parsed = await parseHtml(page, EXCEL_HTML_FORMULA);
    const c1 = getCell(parsed, 0, 2);

    expect(String(c1?.v ?? '')).toBe('3');
    expect(typeof c1?.f).toBe('string');
    expect(String(c1?.f).toUpperCase()).toContain('SUM');
  });

  test('parser helper does not throw on normal Excel HTML', async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    page.on('console', (msg) => {
      if (msg.type() === 'error') errors.push(msg.text());
    });

    const parsed = await parseHtml(page, EXCEL_HTML_BASIC);

    expect(parsed.cells.length).toBeGreaterThan(0);
    const fatal = errors.filter((e) => /Error|TypeError|Uncaught/.test(e));
    expect(fatal).toEqual([]);
  });
});
