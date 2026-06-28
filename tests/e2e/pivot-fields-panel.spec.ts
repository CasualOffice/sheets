import { expect, test, type Page } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * PivotTable Fields task pane — Excel's field list + the four drop zones
 * (Filters / Columns / Rows / Values). Drives the real app:
 *
 *   - The pane auto-opens on insert and reflects the pivot's structure.
 *   - Assigning a field to a zone (the field-list "+" menu) re-applies
 *     the pivot live — no delete-and-reinsert.
 *   - Editing a Values entry's aggregation re-applies live.
 *   - Removing a chip re-applies live.
 *   - The rail / View toggle opens it; with no pivot it shows the empty
 *     state.
 *
 * Reconfiguration is unit-tested in fields-model.unit.test.ts; this
 * asserts the wiring + that the on-sheet cells actually change.
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

/** Insert a Region-rows / Sum-of-Sales pivot at E1. The pane auto-opens. */
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

test.describe('PivotTable Fields pane', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await seedSales(page);
  });

  test('opens on insert and reflects the pivot structure', async ({ page }) => {
    await insertSalesPivot(page);
    await expect(page.getByTestId('pivot-fields-panel')).toBeVisible();

    // Field list carries every source column.
    const list = page.getByTestId('pivot-fields-list');
    await expect(list).toContainText('Region');
    await expect(list).toContainText('Quarter');
    await expect(list).toContainText('Sales');

    // Rows zone shows Region; Values shows "Sum of Sales".
    await expect(page.getByTestId('pivot-fields-zone-rows')).toContainText('Region');
    await expect(page.getByTestId('pivot-fields-zone-values')).toContainText('Sum of Sales');
    // Columns + Filters start empty.
    await expect(page.getByTestId('pivot-fields-zone-cols')).toContainText('Drop or add fields');
  });

  test('adding a field to Columns re-applies the pivot live as a matrix', async ({ page }) => {
    await insertSalesPivot(page);
    // Row-only layout to start: F1 is the value header.
    expect(await readCell(page, 'F1')).toBe('Sum of Sales');

    // Quarter is source column index 1 → open its menu, add to Columns.
    await page.getByTestId('pivot-fields-add-1').click();
    await page.getByTestId('pivot-fields-add-1-cols').click();
    await page.waitForTimeout(150);

    // Now a cross-tab: E1 Region | F1 Q1 | G1 Q2 | H1 Grand Total.
    expect(await readCell(page, 'F1')).toBe('Q1');
    expect(await readCell(page, 'G1')).toBe('Q2');
    expect(await readCell(page, 'H1')).toBe('Grand Total');
    expect(Number(await readCell(page, 'F2'))).toBe(100); // North/Q1
    expect(Number(await readCell(page, 'H2'))).toBe(220); // North total

    // The Columns zone now lists Quarter.
    await expect(page.getByTestId('pivot-fields-zone-cols')).toContainText('Quarter');
  });

  test('changing a Values aggregation re-applies live', async ({ page }) => {
    await insertSalesPivot(page);
    expect(Number(await readCell(page, 'F2'))).toBe(220); // Sum, North

    await page.getByTestId('pivot-fields-values-agg-0').selectOption('average');
    await page.waitForTimeout(150);

    expect(await readCell(page, 'F1')).toBe('Average of Sales');
    expect(Number(await readCell(page, 'F2'))).toBe(110); // (100+120)/2
    expect(Number(await readCell(page, 'F3'))).toBe(87.5); // (80+95)/2
  });

  test('removing a Columns chip reverts to the row-only layout', async ({ page }) => {
    await insertSalesPivot(page);
    await page.getByTestId('pivot-fields-add-1').click();
    await page.getByTestId('pivot-fields-add-1-cols').click();
    await page.waitForTimeout(150);
    expect(await readCell(page, 'F1')).toBe('Q1');

    await page.getByTestId('pivot-fields-chip-cols-0-remove').click();
    await page.waitForTimeout(150);
    expect(await readCell(page, 'F1')).toBe('Sum of Sales');
    expect(Number(await readCell(page, 'F2'))).toBe(220);
  });

  test('rail toggle opens the pane; empty state shows with no pivot', async ({ page }) => {
    // No pivot yet → toggling the rail shows the empty state + CTA.
    await page.getByTestId('panel-rail-pivot').click();
    await expect(page.getByTestId('pivot-fields-panel')).toBeVisible();
    await expect(page.getByTestId('pivot-fields-empty')).toBeVisible();
    await expect(page.getByTestId('pivot-fields-insert-cta')).toBeVisible();
  });
});
