import { expect, test, type Page } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * PivotTable Fields pane — auto-follow the active selection. Clicking into a
 * pivot's output switches the pane to that pivot (Excel's Field List
 * behaviour). With two pivots that have different row fields, activating a
 * cell inside each flips the Rows zone accordingly.
 */

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
}

/** Insert a pivot: rowFieldIndex rows, Sum of Sales, at target. */
async function insertPivot(page: Page, rowFieldIndex: string, target: string) {
  await page.getByTestId('menubar-insert').click();
  await page.getByTestId('menu-item-insert-pivot').click();
  await expect(page.getByTestId('insert-pivot-dialog')).toBeVisible();
  await page.getByTestId('insert-pivot-range').fill('A1:C5');
  await page.getByTestId('insert-pivot-target').fill(target);
  await page.getByTestId('insert-pivot-row-field').selectOption(rowFieldIndex);
  await page.getByTestId('insert-pivot-value-field').selectOption('2');
  await page.getByTestId('insert-pivot-aggregation').selectOption('sum');
  await page.getByTestId('insert-pivot-confirm').click();
  await page.waitForTimeout(150);
}

async function activateCell(page: Page, a1: string) {
  await page.evaluate((ref) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange(ref).activate();
  }, a1);
  await page.waitForTimeout(120);
}

test('clicking into a pivot switches the Fields pane to it', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await seedSales(page);

  // Pivot A: Region rows at E1. Pivot B: Quarter rows at E10.
  await insertPivot(page, '0', 'E1');
  await insertPivot(page, '1', 'E10');

  await expect(page.getByTestId('pivot-fields-panel')).toBeVisible();
  // Two pivots → the picker is shown; the pane defaults to the latest (B).
  await expect(page.getByTestId('pivot-fields-picker')).toBeVisible();
  await expect(page.getByTestId('pivot-fields-zone-rows')).toContainText('Quarter');

  // Activate a cell inside pivot A (E2) → pane follows to the Region pivot.
  await activateCell(page, 'E2');
  await expect(page.getByTestId('pivot-fields-zone-rows')).toContainText('Region');

  // Activate a cell inside pivot B (E11) → pane follows back to Quarter.
  await activateCell(page, 'E11');
  await expect(page.getByTestId('pivot-fields-zone-rows')).toContainText('Quarter');

  // Clicking outside any pivot leaves the pane on the last one (Excel-style).
  await activateCell(page, 'A1');
  await expect(page.getByTestId('pivot-fields-zone-rows')).toContainText('Quarter');
});
