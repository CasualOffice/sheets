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

/**
 * Text-style fidelity: strikethrough + text rotation. Both were dropped by the
 * style mapping. Strikethrough maps to Univer `st: { s: 1 }`; text rotation maps
 * to `tr: { a: <deg> }` (or `{ a: 0, v: 1 }` for Excel's stacked 'vertical').
 * Verifies each survives a full xlsx round-trip.
 */
declare global {
  interface Window {
    __xlsx?: {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      xlsxToWorkbookData: (buf: ArrayBuffer) => Promise<any>;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      workbookDataToXlsx: (data: any) => Promise<Blob>;
    };
  }
}

test('strikethrough + text rotation survive import and round-trip', async ({ page }) => {
  test.setTimeout(60_000);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'struck';
  ws.getCell('A1').font = { strike: true };
  ws.getCell('A2').value = 'tilted';
  ws.getCell('A2').alignment = { textRotation: 45 };
  ws.getCell('A3').value = 'stacked';
  ws.getCell('A3').alignment = { textRotation: 'vertical' };
  const bytes = Array.from(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));

  await page.goto('/');
  await waitForUniver(page);

  const out = await page.evaluate(async (buf: number[]) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const xlsx = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const read = (data: any) => {
      const sheet = data.sheets[data.sheetOrder[0]];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const styleOf = (r: number) => {
        const cell = sheet.cellData?.[r]?.[0];
        return typeof cell?.s === 'string' ? data.styles?.[cell.s] : cell?.s;
      };
      return {
        strike: styleOf(0)?.st?.s ?? null,
        rotA: styleOf(1)?.tr?.a ?? null,
        vertical: styleOf(2)?.tr?.v ?? null,
      };
    };
    const imported = await xlsx.xlsxToWorkbookData(new Uint8Array(buf).buffer);
    const blob = await xlsx.workbookDataToXlsx(imported);
    const round = await xlsx.xlsxToWorkbookData(await blob.arrayBuffer());
    return { imported: read(imported), round: read(round) };
  }, bytes);

  const expected = { strike: 1, rotA: 45, vertical: 1 };
  expect(out.imported).toEqual(expected);
  expect(out.round).toEqual(expected);
});
