import { expect, test, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Sort by colour (Data → Sort by colour…). The partition maths is unit-tested
 * in color-sort.ts; this colours some cells, drives the dialog, and asserts the
 * coloured rows rise to the top carrying their whole row.
 */

async function readCol(page: Page, col: string, n: number) {
  return page.evaluate(
    ({ c, count }) => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      const out: unknown[] = [];
      for (let r = 1; r <= count; r += 1) out.push(ws.getRange(`${c}${r}`).getValue());
      return out;
    },
    { c: col, count: n },
  );
}

test('brings cell-coloured rows to the top, carrying the whole row', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Region' });
    ws.getRange('B1').setValue({ v: 'Sales' });
    const data = [
      ['a', 1],
      ['b', 2],
      ['c', 3],
      ['d', 4],
    ];
    data.forEach((p, i) => {
      ws.getRange('A' + (i + 2)).setValue({ v: p[0] });
      ws.getRange('B' + (i + 2)).setValue({ v: p[1] });
    });
    // Colour the 'b' (A3) and 'd' (A5) key cells red.
    ws.getRange('A3').setBackgroundColor('#ff0000');
    ws.getRange('A5').setBackgroundColor('#ff0000');
    ws.getRange('A1:B5').activate();
  });

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-sort-by-color').click();
  await expect(page.getByTestId('color-sort-dialog')).toBeVisible();

  // Distinct colours in column A (data rows): No-fill first, red second.
  await page.getByTestId('color-sort-swatch-1').click();
  await page.getByTestId('color-sort-ok').click();
  await expect(page.getByTestId('color-sort-dialog')).toBeHidden();

  // Red rows (b, d) rise to the top in order, then the rest (a, c).
  expect(await readCol(page, 'A', 5)).toEqual(['Region', 'b', 'd', 'a', 'c']);
  expect(await readCol(page, 'B', 5)).toEqual(['Sales', 2, 4, 1, 3]);
});

test('sorts by font colour when "Sort on: Font colour" is chosen', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Region' });
    ws.getRange('B1').setValue({ v: 'Sales' });
    const data = [
      ['a', 1],
      ['b', 2],
      ['c', 3],
      ['d', 4],
    ];
    data.forEach((p, i) => {
      ws.getRange('A' + (i + 2)).setValue({ v: p[0] });
      ws.getRange('B' + (i + 2)).setValue({ v: p[1] });
    });
    // Colour the FONT of 'b' (A3) and 'd' (A5) red; backgrounds untouched.
    ws.getRange('A3').setFontColor('#ff0000');
    ws.getRange('A5').setFontColor('#ff0000');
    ws.getRange('A1:B5').activate();
  });

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-sort-by-color').click();
  await page.getByTestId('color-sort-on').selectOption('font');
  // Distinct font colours in column A: Automatic first, red second.
  await page.getByTestId('color-sort-swatch-1').click();
  await page.getByTestId('color-sort-ok').click();
  await expect(page.getByTestId('color-sort-dialog')).toBeHidden();

  expect(await readCol(page, 'A', 5)).toEqual(['Region', 'b', 'd', 'a', 'c']);
  expect(await readCol(page, 'B', 5)).toEqual(['Sales', 2, 4, 1, 3]);
});
