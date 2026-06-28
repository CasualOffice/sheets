import { expect, test, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Watch Window (Formulas → Watch Window). Pins cells and shows their live
 * value + formula; values update as the source changes. The list maths is
 * unit-tested in watch-model.ts; this drives the real panel.
 */

async function activate(page: Page, a1: string) {
  await page.evaluate((ref) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange(ref).activate();
  }, a1);
}

test('add watches, see live values + formula, update, remove', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 10 });
    ws.getRange('B1').setValue({ f: '=A1*2' });
  });

  // Open the panel from the rail; starts empty.
  await page.getByTestId('panel-rail-watch').click();
  await expect(page.getByTestId('watch-panel')).toBeVisible();
  await expect(page.getByTestId('watch-empty')).toBeVisible();

  // Watch A1 (a constant) and B1 (a formula).
  await activate(page, 'A1');
  await page.getByTestId('watch-add').click();
  await activate(page, 'B1');
  await page.getByTestId('watch-add').click();

  await expect(page.getByTestId('watch-value-A1')).toHaveText('10');
  await expect(page.getByTestId('watch-value-B1')).toHaveText('20');
  // The formula column shows B1's formula.
  await expect(page.getByTestId('watch-table')).toContainText('=A1*2');

  // Editing the source updates both watched values live.
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 50 });
  });
  await expect(page.getByTestId('watch-value-A1')).toHaveText('50');
  await expect(page.getByTestId('watch-value-B1')).toHaveText('100');

  // Remove the A1 watch; B1 stays.
  await page.getByTestId('watch-remove-A1').click();
  await expect(page.getByTestId('watch-value-A1')).toHaveCount(0);
  await expect(page.getByTestId('watch-value-B1')).toHaveText('100');
});
