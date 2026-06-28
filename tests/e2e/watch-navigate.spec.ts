import { expect, test, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Watch Window navigation — double-clicking a watch row jumps to that cell
 * (Excel's Watch Window behaviour).
 */

async function activeCell(page: Page) {
  return page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    const r = ws.getSelection().getActiveRange();
    return { row: r.getRow(), col: r.getColumn() };
  });
}

test('double-clicking a watch jumps to its cell', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('C5').setValue({ v: 42 });
    ws.getRange('C5').activate();
  });

  // Watch C5, then move the cursor away.
  await page.getByTestId('panel-rail-watch').click();
  await page.getByTestId('watch-add').click();
  await expect(page.getByTestId('watch-value-C5')).toHaveText('42');
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').activate();
  });
  expect(await activeCell(page)).toEqual({ row: 0, col: 0 });

  // Double-click the watch row → selection jumps back to C5 (row 4, col 2).
  await page.getByTestId('watch-value-C5').dblclick();
  expect(await activeCell(page)).toEqual({ row: 4, col: 2 });
});
