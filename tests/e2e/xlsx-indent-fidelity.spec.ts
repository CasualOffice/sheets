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
 * Cell indentation fidelity. Excel's `alignment.indent` (a level) was dropped on
 * import; it now maps to Univer's left padding `pd.l` (which the renderer applies
 * as a text indent) and round-trips back to the level. Levels 1 and 3 are
 * checked through a full export → re-import.
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

test('cell indentation survives import + round-trip', async ({ page }) => {
  test.setTimeout(60_000);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'flush';
  ws.getCell('A2').value = 'indent1';
  ws.getCell('A2').alignment = { indent: 1 };
  ws.getCell('A3').value = 'indent3';
  ws.getCell('A3').alignment = { indent: 3 };
  const bytes = Array.from(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));

  await page.goto('/');
  await waitForUniver(page);

  const out = await page.evaluate(async (buf: number[]) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const xlsx = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const padOf = (data: any, row: number) => {
      const sheet = data.sheets[data.sheetOrder[0]];
      const cell = sheet.cellData?.[row]?.[0];
      const style = typeof cell?.s === 'string' ? data.styles?.[cell.s] : cell?.s;
      return style?.pd?.l ?? null;
    };
    const imported = await xlsx.xlsxToWorkbookData(new Uint8Array(buf).buffer);
    const blob = await xlsx.workbookDataToXlsx(imported);
    const round = await xlsx.xlsxToWorkbookData(await blob.arrayBuffer());
    return {
      // Row 0 (flush) has no indent → no pd; rows 1/3 get 2 + level*10 px.
      importedPad: [padOf(imported, 0), padOf(imported, 1), padOf(imported, 2)],
      roundPad: [padOf(round, 0), padOf(round, 1), padOf(round, 2)],
    };
  }, bytes);

  // level 0 → no padding; level 1 → 12px; level 3 → 32px. Stable on round-trip.
  expect(out.importedPad).toEqual([null, 12, 32]);
  expect(out.roundPad).toEqual([null, 12, 32]);
});
