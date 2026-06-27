import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Advanced Filter (Data → Advanced Filter…), copy-to-another-location mode.
 * Criteria matching is unit-tested in advanced-filter.ts; this drives the
 * dialog end-to-end: seed a list + a criteria range, run, and check the copied
 * rows land at the destination.
 */

test('copies rows matching the criteria range to a destination', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    // List A1:B5
    const list = [
      ['Region', 'Amount'],
      ['East', 90],
      ['West', 40],
      ['East', 20],
      ['North', 200],
    ];
    list.forEach((row, r) => {
      ws.getRange(r, 0).setValue({ v: row[0] });
      ws.getRange(r, 1).setValue({ v: row[1] });
    });
    // Criteria range D1:E2 — Region = East AND Amount > 50.
    ws.getRange(0, 3).setValue({ v: 'Region' });
    ws.getRange(0, 4).setValue({ v: 'Amount' });
    ws.getRange(1, 3).setValue({ v: 'East' });
    ws.getRange(1, 4).setValue({ v: '>50' });
  });

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-advanced-filter').click();
  await expect(page.getByTestId('advanced-filter-dialog')).toBeVisible();

  await page.getByTestId('advanced-filter-list').fill('A1:B5');
  await page.getByTestId('advanced-filter-criteria').fill('D1:E2');
  await page.getByTestId('advanced-filter-dest').fill('G1');
  await page.getByTestId('advanced-filter-ok').click();

  await expect(page.getByTestId('advanced-filter-result')).toContainText('1 row copied');

  // Destination G1:H2 = header + the single matching row (East / 90).
  const out = await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    const grab = (r: number, c: number) => ws.getRange(r, c).getValue() ?? null;
    return [
      [grab(0, 6), grab(0, 7)],
      [grab(1, 6), grab(1, 7)],
      [grab(2, 6), grab(2, 7)],
    ];
  });
  expect(out[0]).toEqual(['Region', 'Amount']);
  expect(out[1]).toEqual(['East', 90]);
  expect(out[2]).toEqual([null, null]); // nothing else copied
});

test('OR across criteria rows', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    const list = [
      ['Region', 'Amount'],
      ['East', 90],
      ['West', 40],
      ['North', 200],
    ];
    list.forEach((row, r) => {
      ws.getRange(r, 0).setValue({ v: row[0] });
      ws.getRange(r, 1).setValue({ v: row[1] });
    });
    // Criteria: Region = West  OR  Amount >= 200.
    ws.getRange(0, 3).setValue({ v: 'Region' });
    ws.getRange(0, 4).setValue({ v: 'Amount' });
    ws.getRange(1, 3).setValue({ v: '=West' });
    ws.getRange(2, 4).setValue({ v: '>=200' });
  });

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-advanced-filter').click();
  await page.getByTestId('advanced-filter-list').fill('A1:B4');
  await page.getByTestId('advanced-filter-criteria').fill('D1:E3');
  await page.getByTestId('advanced-filter-dest').fill('G1');
  await page.getByTestId('advanced-filter-ok').click();

  await expect(page.getByTestId('advanced-filter-result')).toContainText('2 rows copied');
});
