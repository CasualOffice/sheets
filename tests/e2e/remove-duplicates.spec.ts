import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Remove Duplicates (Data → Remove Duplicates…). Drops rows duplicating an
 * earlier row across the chosen columns, compacts survivors to the top of the
 * selection (preserving value/formula/style), clears the freed rows, and
 * reports the counts. Dedupe maths is unit-tested in dedupe.ts; this drives the
 * dialog end-to-end.
 */

async function seedAndSelect(page: import('@playwright/test').Page, a1: string) {
  await page.evaluate((sel) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    const data: Array<[string, number]> = [
      ['Name', 0], // header
      ['Ada', 90],
      ['Bob', 88],
      ['Ada', 90], // dup of row 1
      ['Cleo', 75],
      ['Bob', 88], // dup of row 2
    ];
    data.forEach((row, r) => {
      ws.getRange(r, 0).setValue({ v: row[0] });
      ws.getRange(r, 1).setValue({ v: r === 0 ? 'Score' : row[1] });
    });
    ws.getRange(sel).activate();
  }, a1);
}

async function colA(page: import('@playwright/test').Page, rows: number) {
  return page.evaluate((n) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    const out: Array<unknown> = [];
    for (let r = 0; r < n; r++) out.push(ws.getRange(r, 0).getCellData()?.v ?? null);
    return out;
  }, rows);
}

test('removes duplicate rows, compacts, and reports the count', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await seedAndSelect(page, 'A1:B6');

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-remove-duplicates').click();
  await expect(page.getByTestId('remove-duplicates-dialog')).toBeVisible();

  // "My data has headers" is on by default; compare all columns.
  await page.getByTestId('remove-duplicates-ok').click();

  // Two duplicate rows removed, three unique remain.
  await expect(page.getByTestId('remove-duplicates-result')).toContainText('2 duplicate');
  await expect(page.getByTestId('remove-duplicates-result')).toContainText('3 unique');

  // Column A is now: header + Ada/Bob/Cleo, then the two freed rows cleared.
  expect(await colA(page, 6)).toEqual(['Name', 'Ada', 'Bob', 'Cleo', null, null]);
});

test('comparing a single column treats same-name rows as duplicates', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await seedAndSelect(page, 'A1:B6');

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-remove-duplicates').click();
  // Uncheck column B (Score) → compare on Name only.
  await page.getByTestId('remove-duplicates-col-1').uncheck();
  await page.getByTestId('remove-duplicates-ok').click();

  // Ada, Bob, Cleo are the only distinct names → 2 removed, 3 remain.
  await expect(page.getByTestId('remove-duplicates-result')).toContainText('2 duplicate');
  expect(await colA(page, 6)).toEqual(['Name', 'Ada', 'Bob', 'Cleo', null, null]);
});

test('reports when there are no duplicates', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ['Name', 'Ada', 'Bob', 'Cleo'].forEach((v, r) => ws.getRange(r, 0).setValue({ v }));
    ws.getRange('A1:A4').activate();
  });

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-remove-duplicates').click();
  await page.getByTestId('remove-duplicates-ok').click();
  await expect(page.getByTestId('remove-duplicates-result')).toContainText(
    'No duplicate values found',
  );
});
