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

/**
 * Drag-and-drop file open. The hook listens on `window` and routes the
 * first supported file through the same loadSpreadsheetFile path File→Open
 * uses, so the assertions mirror xlsx-open.spec.ts.
 *
 * Synthesizing a file drop without Playwright's filechooser API requires
 * building a DataTransfer in page context and dispatching the events
 * manually — see `dropFileOnWindow` below.
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

/**
 * Builds an .xlsx Blob in page context, then dispatches drag* + drop events
 * with a DataTransfer carrying it. Returns once the drop handler has had a
 * chance to start the async load.
 */
async function dropFileOnWindow(page: Page, filename: string, cellValue: string) {
  await page.evaluate(
    async ({ name, v }) => {
      const data = {
        id: `wb-drop-${Date.now()}`,
        rev: 1,
        name: 'drop-fixture',
        appVersion: '0.22.1',
        locale: 1,
        styles: {},
        sheetOrder: ['s1'],
        sheets: {
          s1: {
            id: 's1',
            name: 'Dropped',
            cellData: { 0: { 0: { v } } },
            rowCount: 1024,
            columnCount: 128,
          },
        },
      };
      const blob = await window.__xlsx!.workbookDataToXlsx(data);
      const file = new File([blob], name, {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const dt = new DataTransfer();
      dt.items.add(file);
      const fire = (type: string) =>
        window.dispatchEvent(
          new DragEvent(type, {
            bubbles: true,
            cancelable: true,
            dataTransfer: dt,
          }),
        );
      fire('dragenter');
      fire('dragover');
      fire('drop');
    },
    { name: filename, v: cellValue },
  );
}

test.describe('Drag-and-drop file open', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await exposeConverters(page);
  });

  test('drop overlay appears while a file is over the page', async ({ page }) => {
    await page.evaluate(() => {
      const dt = new DataTransfer();
      dt.items.add(new File(['x'], 'x.xlsx'));
      window.dispatchEvent(
        new DragEvent('dragenter', { bubbles: true, cancelable: true, dataTransfer: dt }),
      );
    });
    await expect(page.getByTestId('file-drop-overlay')).toBeVisible();

    await page.evaluate(() => {
      const dt = new DataTransfer();
      dt.items.add(new File(['x'], 'x.xlsx'));
      window.dispatchEvent(
        new DragEvent('dragleave', { bubbles: true, cancelable: true, dataTransfer: dt }),
      );
    });
    await expect(page.getByTestId('file-drop-overlay')).toBeHidden();
  });

  test('dropping an xlsx file replaces the active workbook', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'BEFORE_DROP' });
    });

    await dropFileOnWindow(page, 'dropped.xlsx', 'DROPPED_CONTENT');

    await page.waitForFunction(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange('A1').getValue() === 'DROPPED_CONTENT';
    }, null, { timeout: 5_000 });

    await expect(page.getByTestId('file-drop-overlay')).toBeHidden();
  });
});
