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
import ExcelJS from 'exceljs';
import { waitForUniver } from './_helpers';

/**
 * Defined names (Excel "named ranges") survive the round-trip through
 * both legs of the pipeline:
 *
 *   - Foreign reader (Microsoft Excel): the workbook's
 *     `<definedName>` part carries the name → range mapping, written by
 *     our exporter from Univer's SHEET_DEFINED_NAME_PLUGIN resource.
 *   - Our own re-open: same plugin resource ships in the hidden
 *     `__casual_sheets_resources__` sidecar and Univer's
 *     DefinedNameDataController hydrates from it.
 *
 * The two tests below pin both directions.
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

async function exposeConverters(page: Page) {
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mod = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    window.__xlsx = mod;
  });
}

test.describe('xlsx defined names round-trip', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await exposeConverters(page);
  });

  test('xlsx authored in Excel keeps its named range after Open in casual-sheets', async ({ page }) => {
    // Build a reference xlsx with a defined name using ExcelJS.
    const refWb = new ExcelJS.Workbook();
    const sheet = refWb.addWorksheet('Data');
    sheet.getCell('A1').value = 10;
    sheet.getCell('A2').value = 20;
    sheet.getCell('A3').value = 30;
    refWb.definedNames.add('Data!$A$1:$A$3', 'Inputs');
    const refBytes = Buffer.from(await refWb.xlsx.writeBuffer());

    // Push through our parser and inspect the snapshot's resources.
    const resources = await page.evaluate(async (bytes) => {
      const buf = new Uint8Array(bytes).buffer;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap: any = await window.__xlsx!.xlsxToWorkbookData(buf);
      return snap.resources as Array<{ name: string; data: string }> | undefined;
    }, Array.from(refBytes));

    const dn = resources?.find((r) => r.name === 'SHEET_DEFINED_NAME_PLUGIN');
    expect(dn, 'SHEET_DEFINED_NAME_PLUGIN resource present').toBeDefined();
    const map = JSON.parse(dn!.data) as Record<string, { name: string; formulaOrRefString: string }>;
    const entry = Object.values(map).find((e) => e.name === 'Inputs');
    expect(entry, 'Inputs named range parsed').toBeDefined();
    expect(entry!.formulaOrRefString).toBe('Data!$A$1:$A$3');
  });

  test('named range we export is visible in xlsx-native defined names', async ({ page }) => {
    // Build a snapshot that already carries the defined-name resource,
    // run our exporter, then parse the output with ExcelJS to confirm
    // it ends up in the workbook's definedNames part.
    const bytes = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap: any = {
        id: 'wb-dn-1',
        rev: 1,
        name: 'wb',
        appVersion: '0.22.1',
        locale: 1,
        styles: {},
        sheetOrder: ['s1'],
        sheets: {
          s1: { id: 's1', name: 'Data', cellData: {}, rowCount: 100, columnCount: 26 },
        },
        resources: [
          {
            name: 'SHEET_DEFINED_NAME_PLUGIN',
            data: JSON.stringify({
              'dn-0': {
                id: 'dn-0',
                name: 'Inputs',
                formulaOrRefString: 'Data!$A$1:$A$3',
              },
            }),
          },
        ],
      };
      const blob = await window.__xlsx!.workbookDataToXlsx(snap);
      return Array.from(new Uint8Array(await blob.arrayBuffer()));
    });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(Buffer.from(bytes));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const model = (wb.definedNames as any).model as Array<{ name: string; ranges: string[] }>;
    const entry = model.find((m) => m.name === 'Inputs');
    expect(entry, 'Inputs defined name present in xlsx').toBeDefined();
    expect(entry!.ranges).toContain('Data!$A$1:$A$3');
  });
});
