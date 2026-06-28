import { expect, test, type Page } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * PivotTable Fields pane — per-pivot refresh. Editing source data does NOT
 * live-update the pivot (Excel refreshes on demand, not on every edit); the
 * pane's Refresh button re-reads the source and recomputes.
 */

async function readCell(page: Page, a1: string) {
  return page.evaluate((ref) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return ws.getRange(ref).getValue();
  }, a1);
}

async function setCell(page: Page, a1: string, v: number) {
  await page.evaluate(
    ({ ref, val }) => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange(ref).setValue({ v: val });
    },
    { ref: a1, val: v },
  );
}

async function seedSales(page: Page) {
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Region' });
    ws.getRange('B1').setValue({ v: 'Quarter' });
    ws.getRange('C1').setValue({ v: 'Sales' });
    ws.getRange('A2').setValue({ v: 'North' });
    ws.getRange('B2').setValue({ v: 'Q1' });
    ws.getRange('C2').setValue({ v: 100 });
    ws.getRange('A3').setValue({ v: 'South' });
    ws.getRange('B3').setValue({ v: 'Q1' });
    ws.getRange('C3').setValue({ v: 80 });
    ws.getRange('A4').setValue({ v: 'North' });
    ws.getRange('B4').setValue({ v: 'Q2' });
    ws.getRange('C4').setValue({ v: 120 });
    ws.getRange('A5').setValue({ v: 'South' });
    ws.getRange('B5').setValue({ v: 'Q2' });
    ws.getRange('C5').setValue({ v: 95 });
    ws.getRange('A1:C5').activate();
  });
  await mainCanvas(page)
    .first()
    .click({ position: { x: 100, y: 100 } });
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1:C5').activate();
  });
}

async function insertSalesPivot(page: Page) {
  await page.getByTestId('menubar-insert').click();
  await page.getByTestId('menu-item-insert-pivot').click();
  await expect(page.getByTestId('insert-pivot-dialog')).toBeVisible();
  await page.getByTestId('insert-pivot-range').fill('A1:C5');
  await page.getByTestId('insert-pivot-target').fill('E1');
  await page.getByTestId('insert-pivot-row-field').selectOption('0');
  await page.getByTestId('insert-pivot-value-field').selectOption('2');
  await page.getByTestId('insert-pivot-aggregation').selectOption('sum');
  await page.getByTestId('insert-pivot-confirm').click();
  await page.waitForTimeout(150);
}

test('editing source leaves the pivot stale until Refresh', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await seedSales(page);
  await insertSalesPivot(page);

  // North = C2(100) + C4(120) = 220.
  expect(Number(await readCell(page, 'F2'))).toBe(220);

  // Edit a source cell. The pivot does NOT live-update (Excel parity).
  await setCell(page, 'C2', 500);
  await page.waitForTimeout(150);
  expect(Number(await readCell(page, 'F2'))).toBe(220);

  // Refresh re-reads the source: North = 500 + 120 = 620.
  await page.getByTestId('pivot-fields-refresh').click();
  await page.waitForTimeout(150);
  expect(Number(await readCell(page, 'F2'))).toBe(620);
});
