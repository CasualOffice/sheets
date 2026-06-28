import { expect, test, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Pivots — nested column fields end-to-end. Region (row) × Quarter × Month
 * (two nested column fields) → two-level column headers and tuple-sliced
 * values. The compute is unit-tested in compute.unit.test.ts; this drives the
 * dialog's "Add column field (nest)" flow.
 */

async function cell(page: Page, a1: string): Promise<unknown> {
  return page.evaluate((c) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return ws.getRange(c).getCellData()?.v;
  }, a1);
}

test('nesting Quarter × Month produces two-level column headers', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    const rows = [
      ['Region', 'Quarter', 'Month', 'Sales'],
      ['East', 'Q1', 'Jan', 10],
      ['East', 'Q1', 'Feb', 20],
      ['East', 'Q2', 'Apr', 30],
      ['West', 'Q1', 'Jan', 40],
    ];
    rows.forEach((r, ri) =>
      r.forEach((v, ci) => ws.getRange(ri, ci).setValue({ v: v as string | number })),
    );
    ws.getRange('A1:D5').activate();
  });

  await page.getByTestId('menubar-insert').click();
  await page.getByTestId('menu-item-insert-pivot').click();
  await expect(page.getByTestId('insert-pivot-dialog')).toBeVisible();

  await page.getByTestId('insert-pivot-range').fill('A1:D5');
  await page.getByTestId('insert-pivot-target').fill('E1');
  await page.getByTestId('insert-pivot-row-field').selectOption('0'); // Region
  await page.getByTestId('insert-pivot-col-field').selectOption('1'); // Quarter
  await page.getByTestId('insert-pivot-col-add').click();
  await page.getByTestId('insert-pivot-col-field-1').selectOption('2'); // Month
  await page.getByTestId('insert-pivot-value-field').selectOption('3'); // Sales
  await page.getByTestId('insert-pivot-aggregation').selectOption('sum');
  await page.getByTestId('insert-pivot-confirm').click();
  await page.waitForTimeout(200);

  // Tuples sort (Q1,Feb) < (Q1,Jan) < (Q2,Apr).
  // Row 1: Quarter span | Row 2: Month level.
  expect(await cell(page, 'E1')).toBe('Region');
  expect(await cell(page, 'F1')).toBe('Q1');
  expect(await cell(page, 'G1')).toBe(''); // span blank
  expect(await cell(page, 'H1')).toBe('Q2');
  expect(await cell(page, 'I1')).toBe('Grand Total');
  expect(await cell(page, 'F2')).toBe('Feb');
  expect(await cell(page, 'G2')).toBe('Jan');
  expect(await cell(page, 'H2')).toBe('Apr');
  // East: Feb 20, Jan 10, Apr 30, total 60.
  expect(await cell(page, 'E3')).toBe('East');
  expect(Number(await cell(page, 'F3'))).toBe(20);
  expect(Number(await cell(page, 'G3'))).toBe(10);
  expect(Number(await cell(page, 'H3'))).toBe(30);
  expect(Number(await cell(page, 'I3'))).toBe(60);
  // West: only Q1/Jan 40.
  expect(Number(await cell(page, 'G4'))).toBe(40);
  expect(Number(await cell(page, 'I4'))).toBe(40);
  // Grand total row.
  expect(await cell(page, 'E5')).toBe('Grand Total');
  expect(Number(await cell(page, 'I5'))).toBe(100);
});
