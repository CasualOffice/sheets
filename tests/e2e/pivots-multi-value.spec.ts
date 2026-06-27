import { expect, test, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Pivots — multiple value fields. The compute already fans out
 * `model.values` into one output column each; this exercises the dialog's
 * value-field list (Add value field) end-to-end: a Region pivot with both
 * Sum of Sales and Count of Sales.
 */

async function seed(page: Page) {
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    const rows = [
      ['Region', 'Product', 'Sales'],
      ['East', 'A', 100],
      ['East', 'B', 200],
      ['East', 'A', 50],
      ['West', 'A', 150],
    ];
    rows.forEach((r, ri) =>
      r.forEach((v, ci) => ws.getRange(ri, ci).setValue({ v: v as string | number })),
    );
    ws.getRange('A1:C5').activate();
  });
}

async function cell(page: Page, a1: string): Promise<unknown> {
  return page.evaluate((c) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return ws.getRange(c).getCellData()?.v;
  }, a1);
}

test('a Region pivot with Sum + Count of Sales emits two value columns', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await seed(page);

  await page.getByTestId('menubar-insert').click();
  await page.getByTestId('menu-item-insert-pivot').click();
  await expect(page.getByTestId('insert-pivot-dialog')).toBeVisible();

  await page.getByTestId('insert-pivot-range').fill('A1:C5');
  await page.getByTestId('insert-pivot-target').fill('E1');
  await page.getByTestId('insert-pivot-row-field').selectOption('0'); // Region

  // First value field: Sum of Sales (col 2).
  await page.getByTestId('insert-pivot-value-field').selectOption('2');
  await page.getByTestId('insert-pivot-aggregation').selectOption('sum');
  // Add a second value field: Count of Sales.
  await page.getByTestId('insert-pivot-value-add').click();
  await page.getByTestId('insert-pivot-value-field-1').selectOption('2');
  await page.getByTestId('insert-pivot-aggregation-1').selectOption('count');

  await page.getByTestId('insert-pivot-confirm').click();
  await page.waitForTimeout(200);

  // Header: Region | Sum of Sales | Count of Sales.
  expect(await cell(page, 'E1')).toBe('Region');
  expect(await cell(page, 'F1')).toBe('Sum of Sales');
  expect(await cell(page, 'G1')).toBe('Count of Sales');
  // East: sum 350, count 3.
  expect(await cell(page, 'E2')).toBe('East');
  expect(Number(await cell(page, 'F2'))).toBe(350);
  expect(Number(await cell(page, 'G2'))).toBe(3);
  // West: sum 150, count 1.
  expect(await cell(page, 'E3')).toBe('West');
  expect(Number(await cell(page, 'F3'))).toBe(150);
  expect(Number(await cell(page, 'G3'))).toBe(1);
  // Grand total: 500 / 4.
  expect(await cell(page, 'E4')).toBe('Grand Total');
  expect(Number(await cell(page, 'F4'))).toBe(500);
  expect(Number(await cell(page, 'G4'))).toBe(4);
});
