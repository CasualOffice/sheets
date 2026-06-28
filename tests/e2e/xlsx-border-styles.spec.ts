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
 * Border-style fidelity. The importer previously hardcoded every border to a
 * thin line, so dashed / double / thick / medium / hair / dotted all collapsed
 * to thin on open (and on export). This verifies each Excel line style maps to
 * its Univer BorderStyleTypes value and survives a full xlsx round-trip.
 *
 * Univer BorderStyleTypes: THIN=1, HAIR=2, DOTTED=3, DASHED=4, DASH_DOT=5,
 * DASH_DOT_DOT=6, DOUBLE=7, MEDIUM=8, MEDIUM_DASHED=9, MEDIUM_DASH_DOT=10,
 * MEDIUM_DASH_DOT_DOT=11, SLANT_DASH_DOT=12, THICK=13.
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

// Excel style string → expected Univer BorderStyleTypes enum value.
const CASES: Array<{ row: number; excel: ExcelJS.BorderStyle; univer: number }> = [
  { row: 1, excel: 'thin', univer: 1 },
  { row: 2, excel: 'dashed', univer: 4 },
  { row: 3, excel: 'double', univer: 7 },
  { row: 4, excel: 'medium', univer: 8 },
  { row: 5, excel: 'thick', univer: 13 },
  { row: 6, excel: 'dotted', univer: 3 },
];

test('xlsx border line styles survive import + round-trip (not all thin)', async ({ page }) => {
  test.setTimeout(60_000);

  // Build a fixture: one cell per border style, each with a top border.
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (const c of CASES) {
    const cell = ws.getCell(`A${c.row}`);
    cell.value = c.excel;
    cell.border = { top: { style: c.excel, color: { argb: 'FF0000FF' } } };
  }
  const bytes = Array.from(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));

  await page.goto('/');
  await waitForUniver(page);

  // Import once, then round-trip (export → re-import) and read both.
  const result = await page.evaluate(async (buf: number[]) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const xlsx = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    const ab = new Uint8Array(buf).buffer;
    const topBorders = (data: {
      sheetOrder: string[];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      sheets: Record<string, any>;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      styles: Record<string, any>;
    }) => {
      const sheet = data.sheets[data.sheetOrder[0]];
      return (row: number) => {
        const cell = sheet.cellData?.[row]?.[0];
        const style = typeof cell?.s === 'string' ? data.styles?.[cell.s] : cell?.s;
        return style?.bd?.t?.s ?? null;
      };
    };

    const imported = await xlsx.xlsxToWorkbookData(ab);
    const blob = await xlsx.workbookDataToXlsx(imported);
    const roundTripped = await xlsx.xlsxToWorkbookData(await blob.arrayBuffer());
    const a = topBorders(imported);
    const b = topBorders(roundTripped);
    const rows = [1, 2, 3, 4, 5, 6];
    return {
      imported: rows.map((r) => a(r - 1)),
      roundTripped: rows.map((r) => b(r - 1)),
    };
  }, bytes);

  const expected = CASES.map((c) => c.univer);
  expect(result.imported).toEqual(expected);
  expect(result.roundTripped).toEqual(expected);
  // Sanity: not everything is thin (the old bug).
  expect(new Set(result.imported).size).toBeGreaterThan(1);
});
