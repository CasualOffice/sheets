import { expect, test } from '@playwright/test';
import { selectRange, waitForUniver } from './_helpers';

/**
 * Data → Show all rows. Recovery action for the case where a filter (sheet- or
 * table-level) left rows hidden after the filter UI was dismissed and there's
 * no obvious surface to unhide them from.
 */
test('Data → Show all rows reveals every hidden row', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  await selectRange(page, 'A1');
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    for (let r = 0; r < 6; r++) ws.getRange(r, 0).setValue({ v: `r${r + 1}` });
    // Hide rows 2..4 (0-indexed: 1..3).
    ws.hideRows(1, 3);
  });

  // FWorksheet doesn't expose isRowHidden; reach through to the underlying
  // Worksheet, which has getRowVisible(row).
  const hiddenBefore = await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fws: any = api.getActiveWorkbook()!.getActiveSheet();
    const ws = fws.getSheet();
    return [1, 2, 3].map((r) => !ws.getRowVisible(r));
  });
  expect(hiddenBefore).toEqual([true, true, true]);

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-show-all-rows').click();

  const hiddenAfter = await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fws: any = api.getActiveWorkbook()!.getActiveSheet();
    const ws = fws.getSheet();
    return [1, 2, 3].map((r) => !ws.getRowVisible(r));
  });
  expect(hiddenAfter).toEqual([false, false, false]);
});
