import { expect, test, type Page } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * PivotTable Fields pane — report filter (Filters zone) slice 2. Placing a
 * field in Filters and unchecking values actually restricts the source
 * records and re-applies the pivot live. The toggle maths is unit-tested in
 * fields-model.unit.test.ts; this drives the real pane + asserts the on-sheet
 * cells change.
 */

async function readCell(page: Page, a1: string) {
  return page.evaluate((ref) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return ws.getRange(ref).getValue();
  }, a1);
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
  await page.getByTestId('insert-pivot-row-field').selectOption('0'); // Region
  await page.getByTestId('insert-pivot-value-field').selectOption('2'); // Sales
  await page.getByTestId('insert-pivot-aggregation').selectOption('sum');
  await page.getByTestId('insert-pivot-confirm').click();
  await page.waitForTimeout(150);
}

test.describe('PivotTable Fields pane — report filter', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await seedSales(page);
    await insertSalesPivot(page);
  });

  test('unchecking a filter value re-applies the pivot live', async ({ page }) => {
    // Region rows / Sum of Sales: North = 100+120 = 220, total = 395.
    expect(Number(await readCell(page, 'F2'))).toBe(220);
    expect(Number(await readCell(page, 'F4'))).toBe(395);

    // Add Quarter (source column 1) to the Filters zone.
    await page.getByTestId('pivot-fields-add-1').click();
    await page.getByTestId('pivot-fields-add-1-filters').click();
    await page.waitForTimeout(100);
    await expect(page.getByTestId('pivot-fields-zone-filters')).toContainText('Quarter');

    // Expand the value checklist and drop Q2 — only Q1 records remain.
    await page.getByTestId('pivot-fields-filter-toggle-0').click();
    await expect(page.getByTestId('pivot-fields-filter-values-0')).toBeVisible();
    await page.getByTestId('pivot-fields-filter-0-Q2').uncheck();
    await page.waitForTimeout(150);

    // North → 100 (Q1 only), South → 80, Grand Total → 180.
    expect(Number(await readCell(page, 'F2'))).toBe(100);
    expect(Number(await readCell(page, 'F3'))).toBe(80);
    expect(Number(await readCell(page, 'F4'))).toBe(180);

    // The chip count reflects the narrowed selection.
    await expect(page.getByTestId('pivot-fields-filter-toggle-0')).toContainText('1 of 2 selected');
  });

  test('Select all restores every record after a narrow', async ({ page }) => {
    await page.getByTestId('pivot-fields-add-1').click();
    await page.getByTestId('pivot-fields-add-1-filters').click();
    await page.getByTestId('pivot-fields-filter-toggle-0').click();
    await page.getByTestId('pivot-fields-filter-0-Q2').uncheck();
    await page.waitForTimeout(150);
    expect(Number(await readCell(page, 'F4'))).toBe(180);

    await page.getByTestId('pivot-fields-filter-all-0').click();
    await page.waitForTimeout(150);
    expect(Number(await readCell(page, 'F2'))).toBe(220);
    expect(Number(await readCell(page, 'F4'))).toBe(395);
    await expect(page.getByTestId('pivot-fields-filter-toggle-0')).toContainText('2 of 2 selected');
  });
});
