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

import { expect, test } from '@playwright/test';
import ExcelJS from 'exceljs';
import { waitForUniver } from './_helpers';

declare global {
  interface Window {
    __xlsx?: {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      xlsxToWorkbookData: (buf: ArrayBuffer) => Promise<any>;
    };
  }
}

/**
 * Regression for #152 ("formula not working"): a workbook with date cells
 * feeding `=NETWORKDAYS(E8,F8)`. ExcelJS surfaces date-formatted cells as JS
 * Dates; importing them as ISO strings (the old behaviour) left date functions
 * unable to parse their operands, so the formula errored. Dates must import as
 * Excel serial numbers (which the formula engine and the date number-format
 * both understand).
 */
test('date cells import as Excel serials so date formulas evaluate', async ({ page }) => {
  // Build the fixture in Node with real Date-valued cells + a NETWORKDAYS formula.
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('E8').value = new Date('2026-06-01T00:00:00Z');
  ws.getCell('E8').numFmt = 'yyyy-mm-dd';
  ws.getCell('F8').value = new Date('2026-06-30T00:00:00Z');
  ws.getCell('F8').numFmt = 'yyyy-mm-dd';
  ws.getCell('G8').value = { formula: 'NETWORKDAYS(E8,F8)', result: 22 } as ExcelJS.CellValue;
  const bytes = Array.from(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));

  await page.goto('/');
  await waitForUniver(page);

  const imported = await page.evaluate(async (buf: number[]) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const xlsx = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    const data = await xlsx.xlsxToWorkbookData(new Uint8Array(buf).buffer);
    const sheet = data.sheets[Object.keys(data.sheets)[0]];
    const cell = (r: number, c: number) => sheet.cellData?.[r]?.[c];
    const e8 = cell(7, 4);
    const styleObj = typeof e8?.s === 'string' ? data.styles?.[e8.s] : e8?.s;
    return {
      e8: e8?.v,
      f8: cell(7, 5)?.v,
      e8Pattern: styleObj?.n?.pattern,
      g8Formula: cell(7, 6)?.f,
    };
  }, bytes);

  // 2026-06-01 → serial 46174, 2026-06-30 → serial 46203 (Excel 1900 epoch).
  expect(imported.e8).toBe(46174);
  expect(imported.f8).toBe(46203);
  // The date number-format survives so the serial still renders as a date.
  expect(imported.e8Pattern).toBeTruthy();
  // The formula is preserved (not flattened to its cached result).
  expect(imported.g8Formula).toBe('=NETWORKDAYS(E8,F8)');
});
