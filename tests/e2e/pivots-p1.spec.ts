import { expect, test, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Pivots P1 — filter fields + refresh.
 *
 * P0 covered single-field group-by aggregation. P1 layers on:
 *
 *   - **Filter fields**: limit source records before bucketing. Verify
 *     the totals reflect only the allowed values.
 *   - **Refresh**: after editing the source range, the Data → Refresh
 *     PivotTables menu re-runs the compute and updates the output.
 *
 * Multi-row-field and drill-down stay out of scope and are noted in
 * the docs/CO-EDITING and CLAUDE files as "later".
 */

async function seedRegionalSales(page: Page) {
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    // Headers
    ws.getRange('A1').setValue({ v: 'Region' });
    ws.getRange('B1').setValue({ v: 'Quarter' });
    ws.getRange('C1').setValue({ v: 'Sales' });
    // Records
    const rows: Array<[string, string, number]> = [
      ['North', 'Q1', 100],
      ['South', 'Q1', 200],
      ['North', 'Q2', 150],
      ['South', 'Q2', 250],
      ['West', 'Q1', 50],
    ];
    rows.forEach(([region, qtr, sales], i) => {
      ws.getRange(`A${i + 2}`).setValue({ v: region });
      ws.getRange(`B${i + 2}`).setValue({ v: qtr });
      ws.getRange(`C${i + 2}`).setValue({ v: sales });
    });
  });
}

async function openInsertPivot(page: Page) {
  await page.getByTestId('menubar-insert').click();
  await page.getByTestId('menu-item-insert-pivot').click();
  await expect(page.getByTestId('insert-pivot-dialog')).toBeVisible();
}

async function readCellValue(page: Page, a1: string): Promise<unknown> {
  return page.evaluate((cell) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return ws.getRange(cell).getCellData()?.v;
  }, a1);
}

test.describe('Pivots P1', () => {
  test('filter field restricts records before aggregation', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await seedRegionalSales(page);
    await openInsertPivot(page);

    await page.getByTestId('insert-pivot-range').fill('A1:C6');
    await page.getByTestId('insert-pivot-target').fill('E1');
    // Row field: Region (col 0), Value field: Sales (col 2), Sum.
    await page.getByTestId('insert-pivot-row-field').selectOption('0');
    await page.getByTestId('insert-pivot-value-field').selectOption('2');
    await page.getByTestId('insert-pivot-aggregation').selectOption('sum');
    // Filter field: Quarter (col 1). Uncheck Q2.
    await page.getByTestId('insert-pivot-filter-field').selectOption('1');
    await page.getByTestId('insert-pivot-filter-Q2').uncheck();
    await page.getByTestId('insert-pivot-confirm').click();
    await page.waitForTimeout(150);

    // With Q2 filtered out, only Q1 rows feed the aggregation.
    // North/Q1=100, South/Q1=200, West/Q1=50 → grand total = 350.
    // Header row at E1, row keys + Grand Total at the bottom.
    // Layout: E1 = "Region", F1 = "Sum of Sales".
    expect(await readCellValue(page, 'E1')).toBe('Region');
    expect(await readCellValue(page, 'F1')).toBe('Sum of Sales');
    // Find the Grand Total row by scanning E2..E10 (cap is 5+grand=6).
    let grandRow = -1;
    for (let r = 2; r <= 10; r += 1) {
      const v = await readCellValue(page, `E${r}`);
      if (v === 'Grand Total') {
        grandRow = r;
        break;
      }
    }
    expect(grandRow).toBeGreaterThan(1);
    const grandSum = await readCellValue(page, `F${grandRow}`);
    expect(Number(grandSum)).toBe(350);
  });

  test('Refresh PivotTables re-computes after a source edit', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await seedRegionalSales(page);
    await openInsertPivot(page);

    await page.getByTestId('insert-pivot-range').fill('A1:C6');
    await page.getByTestId('insert-pivot-target').fill('E1');
    await page.getByTestId('insert-pivot-row-field').selectOption('0');
    await page.getByTestId('insert-pivot-value-field').selectOption('2');
    await page.getByTestId('insert-pivot-aggregation').selectOption('sum');
    await page.getByTestId('insert-pivot-confirm').click();
    await page.waitForTimeout(150);

    // Initial grand total = 100+200+150+250+50 = 750. Find + assert.
    const findGrandSum = async (): Promise<number> => {
      for (let r = 2; r <= 10; r += 1) {
        const v = await readCellValue(page, `E${r}`);
        if (v === 'Grand Total') {
          const s = await readCellValue(page, `F${r}`);
          return Number(s);
        }
      }
      return NaN;
    };
    expect(await findGrandSum()).toBe(750);

    // Bump North/Q1 by 1000 in the source and confirm the pivot still
    // shows 750 (stale). Refresh, then it should be 1750.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('C2').setValue({ v: 1100 }); // was 100
    });
    expect(await findGrandSum()).toBe(750);

    await page.getByTestId('menubar-data').click();
    await page.getByTestId('menu-item-refresh-pivots').click();
    await page.waitForTimeout(200);
    expect(await findGrandSum()).toBe(1750);
  });
});
