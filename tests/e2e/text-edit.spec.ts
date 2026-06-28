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
import { mainCanvas, readCell, selectRange, waitForUniver } from './_helpers';

/**
 * Text-editing audit — confirms the core keyboard flows users expect from
 * Excel actually work through the canvas + Univer.
 */

import type { Page } from '@playwright/test';

async function clickCellAt(page: Page, offsetX: number, offsetY: number) {
  const grid = mainCanvas(page);
  const box = await grid.boundingBox();
  if (!box) throw new Error('grid not ready');
  await page.mouse.click(box.x + offsetX, box.y + offsetY);
}

test.describe('Text editing', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('F2 opens the inline editor on the active cell', async ({ page }) => {
    // Seed a value, select the cell via canvas click, press F2 to start editing.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('B2').setValue({ v: 'seed' });
    });
    await clickCellAt(page, 200, 60);

    await page.keyboard.press('F2');
    // Append to the existing value
    await page.keyboard.type('-edit');
    await page.keyboard.press('Enter');

    const cell = await readCell(page, 'B2');
    expect(cell?.v).toBe('seed-edit');
  });

  test('Delete clears the active cell', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('B2').setValue({ v: 'to-delete' });
    });
    await clickCellAt(page, 200, 60);

    await page.keyboard.press('Delete');

    const cell = await readCell(page, 'B2');
    expect(cell?.v == null || cell?.v === '').toBe(true);
  });

  test('Backspace starts editing with an empty cell', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('B2').setValue({ v: 'wiped' });
    });
    await clickCellAt(page, 200, 60);

    await page.keyboard.press('Backspace');
    await page.keyboard.type('fresh');
    await page.keyboard.press('Enter');

    const cell = await readCell(page, 'B2');
    expect(cell?.v).toBe('fresh');
  });

  test('Escape during edit cancels the typed text', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('B2').setValue({ v: 'keep-me' });
    });
    await clickCellAt(page, 200, 60);

    await page.keyboard.press('F2');
    await page.keyboard.type('-junk');
    await page.keyboard.press('Escape');

    const cell = await readCell(page, 'B2');
    expect(cell?.v).toBe('keep-me');
  });
});

test.describe('Cell merge', () => {
  test('Merge & Center merges the selection and centers it', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'Title' });
    });
    await selectRange(page, 'A1:C1');

    const merge = page.getByTestId('ribbon-btn-merge-cells');
    await expect(merge).toBeEnabled();
    await merge.click();

    const merges = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getMergedRanges().map(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (r: any) => ({
          startRow: r.getRange().startRow,
          startColumn: r.getRange().startColumn,
          endRow: r.getRange().endRow,
          endColumn: r.getRange().endColumn,
        }),
      );
    });
    expect(merges).toEqual([{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 2 }]);

    // Re-selecting the merged range shows the button as pressed.
    await selectRange(page, 'A1:C1');
    await expect(merge).toHaveAttribute('aria-pressed', 'true');
  });

  test('Pressing merge again unmerges', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1:B2').merge();
    });
    await selectRange(page, 'A1:B2');

    const merge = page.getByTestId('ribbon-btn-merge-cells');
    await expect(merge).toHaveAttribute('aria-pressed', 'true');
    await merge.click();

    const count = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getMergedRanges().length;
    });
    expect(count).toBe(0);
  });

  test('Merge button is disabled for a single-cell selection', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await selectRange(page, 'A1');
    await expect(page.getByTestId('ribbon-btn-merge-cells')).toBeDisabled();
  });
});

test.describe('Workbook growth', () => {
  test('Initial size is 1024 × 26', async ({ page }) => {
    // 26 columns = A..Z. Univer allocates row/column metadata up-front
    // for the declared count; useWorkbookGrowth extends it on demand.
    await page.goto('/');
    await waitForUniver(page);

    const size = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return { rows: ws.getMaxRows(), cols: ws.getMaxColumns() };
    });
    expect(size).toEqual({ rows: 1024, cols: 26 });
  });

  test('Selecting near the bottom edge grows rows', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    // Select a cell in row 1000 (0-indexed: 999), which is within the 32-row
    // edge buffer of the initial 1024-row sheet.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange(999, 0).activate();
    });

    const after = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getMaxRows();
    });
    expect(after).toBeGreaterThan(1024);
    expect(after).toBeLessThanOrEqual(8192);
  });

  test('Selecting near the right edge grows columns', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    // Column 22 is within the 8-column edge buffer of the initial 26.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange(0, 22).activate();
    });

    const after = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getMaxColumns();
    });
    expect(after).toBeGreaterThan(26);
    expect(after).toBeLessThanOrEqual(1024);
  });
});
