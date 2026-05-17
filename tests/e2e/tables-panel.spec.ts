import { expect, test } from '@playwright/test';
import { selectRange, waitForUniver } from './_helpers';

test.describe('Tables panel', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('opens from Data menu and shows the empty state initially', async ({ page }) => {
    await expect(page.getByTestId('tables-panel')).toHaveCount(0);
    await page.getByTestId('menubar-data').click();
    await page.getByTestId('menu-item-tables-panel').click();
    await expect(page.getByTestId('tables-panel')).toBeVisible();
    await expect(page.getByTestId('tables-panel')).toContainText(/no tables on this sheet/i);
  });

  test('lists a table after Format-as-Table is applied', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'Name' });
      ws.getRange('A2').setValue({ v: 'Ada' });
      ws.getRange('B1').setValue({ v: 'Score' });
      ws.getRange('B2').setValue({ v: 92 });
    });
    await selectRange(page, 'A1:B2');
    await page.getByTestId('ribbon-dropdown-format-as-table-apply').click();
    await page.waitForTimeout(600);

    await page.getByTestId('menubar-data').click();
    await page.getByTestId('menu-item-tables-panel').click();

    const panel = page.getByTestId('tables-panel');
    await expect(panel).toBeVisible();
    // One row, range rendered in A1 notation.
    await expect(panel).toContainText('A1:B2');
  });
});

test('growth hook skips Select-All so corner click doesn\'t freeze', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  // Programmatically select the entire sheet (the same shape the corner
  // click produces). The growth hook would otherwise grow the sheet in a
  // loop until MAX_ROWS, freezing for ~800ms.
  const before = await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fws: any = api.getActiveWorkbook()!.getActiveSheet();
    return { rows: fws.getMaxRows(), cols: fws.getMaxColumns() };
  });

  const t0 = Date.now();
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fws: any = api.getActiveWorkbook()!.getActiveSheet();
    fws.getRange(0, 0, fws.getMaxRows(), fws.getMaxColumns()).activate();
  });
  const dur = Date.now() - t0;

  const after = await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fws: any = api.getActiveWorkbook()!.getActiveSheet();
    return { rows: fws.getMaxRows(), cols: fws.getMaxColumns() };
  });

  // Behavior assertion: the sheet must NOT have grown. The pre-fix loop
  // chased the edge in 256-row chunks all the way to MAX_ROWS, so a
  // Select-All would have ended at rows=MAX_ROWS and cols=MAX_COLUMNS. With
  // the guard in place the dimensions are unchanged.
  expect(after.rows).toBe(before.rows);
  expect(after.cols).toBe(before.cols);
  // (dur captured for local debugging; not asserted — CI runners are too
  // variable to make a perf threshold meaningful.)
  void dur;
});
