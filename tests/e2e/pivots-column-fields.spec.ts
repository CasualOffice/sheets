import { expect, test, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Pivots P2 — column fields (cross-tab / matrix layout).
 *
 * Placing a field on the COLUMN axis fans the value field out across
 * one column per distinct column-field value, producing Excel's
 * cross-tab layout:
 *
 *     [ Region   | Q1  | Q2  | Grand Total ]
 *     [ North    | 100 | 120 |         220 ]
 *     [ South    |  80 |  95 |         175 ]
 *     [ Grand T. | 180 | 215 |         395 ]
 *
 * The Insert dialog gains a "Column field" dropdown. The bottom Grand
 * Total row carries column totals; the right-hand Grand Total column
 * carries row totals; the bottom-right cell is the overall total.
 */

async function seedRegionQuarterSales(page: Page) {
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Region' });
    ws.getRange('B1').setValue({ v: 'Quarter' });
    ws.getRange('C1').setValue({ v: 'Sales' });
    const rows: Array<[string, string, number]> = [
      ['North', 'Q1', 100],
      ['South', 'Q1', 80],
      ['North', 'Q2', 120],
      ['South', 'Q2', 95],
    ];
    rows.forEach(([region, quarter, sales], i) => {
      ws.getRange(`A${i + 2}`).setValue({ v: region });
      ws.getRange(`B${i + 2}`).setValue({ v: quarter });
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

test.describe('Pivots — column fields (cross-tab matrix)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await seedRegionQuarterSales(page);
    await openInsertPivot(page);
  });

  test('Region rows × Quarter columns produces a matrix with row + column totals', async ({
    page,
  }) => {
    await page.getByTestId('insert-pivot-range').fill('A1:C5');
    await page.getByTestId('insert-pivot-target').fill('E1');
    // Row field = Region (col 0), Column field = Quarter (col 1),
    // Value field = Sales (col 2), Sum.
    await page.getByTestId('insert-pivot-row-field').selectOption('0');
    await page.getByTestId('insert-pivot-col-field').selectOption('1');
    await page.getByTestId('insert-pivot-value-field').selectOption('2');
    await page.getByTestId('insert-pivot-aggregation').selectOption('sum');
    await page.getByTestId('insert-pivot-confirm').click();
    await page.waitForTimeout(150);

    // Expected matrix, starting at E1:
    //   E1 = "Region"   F1 = "Q1"   G1 = "Q2"   H1 = "Grand Total"
    //   E2 = "North"    F2 = 100    G2 = 120    H2 = 220
    //   E3 = "South"    F3 = 80     G3 = 95     H3 = 175
    //   E4 = "Grand T." F4 = 180    G4 = 215    H4 = 395
    expect(await readCellValue(page, 'E1')).toBe('Region');
    expect(await readCellValue(page, 'F1')).toBe('Q1');
    expect(await readCellValue(page, 'G1')).toBe('Q2');
    expect(await readCellValue(page, 'H1')).toBe('Grand Total');

    expect(await readCellValue(page, 'E2')).toBe('North');
    expect(Number(await readCellValue(page, 'F2'))).toBe(100);
    expect(Number(await readCellValue(page, 'G2'))).toBe(120);
    expect(Number(await readCellValue(page, 'H2'))).toBe(220);

    expect(await readCellValue(page, 'E3')).toBe('South');
    expect(Number(await readCellValue(page, 'F3'))).toBe(80);
    expect(Number(await readCellValue(page, 'G3'))).toBe(95);
    expect(Number(await readCellValue(page, 'H3'))).toBe(175);

    expect(await readCellValue(page, 'E4')).toBe('Grand Total');
    expect(Number(await readCellValue(page, 'F4'))).toBe(180);
    expect(Number(await readCellValue(page, 'G4'))).toBe(215);
    expect(Number(await readCellValue(page, 'H4'))).toBe(395);
  });

  test('a column-field PivotModel round-trips through the resource payload', async ({ page }) => {
    // Prove the model schema + validator accept a populated `cols` list
    // and that it survives the resources write/read cycle (which is how
    // pivots travel through xlsx + collab). Uses the real resource
    // helpers rather than poking React state.
    const cols = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snapshot: any = wb.save();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const res = await import(/* @vite-ignore */ '/src/pivots/resources.ts' as any);
      const model = {
        id: 'pt-matrix',
        sourceSheetId: snapshot.sheetOrder[0],
        source: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 },
        targetSheetId: snapshot.sheetOrder[0],
        target: { row: 0, column: 4 },
        rows: [{ column: 0 }],
        cols: [{ column: 1 }],
        values: [{ column: 2, agg: 'sum' }],
        title: 'PivotTable 1',
      };
      res.writePivotsIntoSnapshot(snapshot, [model]);
      const reloaded = res.readPivotsFromSnapshot(snapshot);
      return reloaded[0]?.cols;
    });
    expect(cols).toEqual([{ column: 1 }]);
  });

  test('choosing the same field for rows and columns falls back to a row-only pivot', async ({
    page,
  }) => {
    // Degenerate guard: Region on both axes shouldn't produce an empty
    // matrix — it collapses to the classic single-column layout.
    await page.getByTestId('insert-pivot-range').fill('A1:C5');
    await page.getByTestId('insert-pivot-target').fill('E1');
    await page.getByTestId('insert-pivot-row-field').selectOption('0');
    await page.getByTestId('insert-pivot-col-field').selectOption('0');
    await page.getByTestId('insert-pivot-value-field').selectOption('2');
    await page.getByTestId('insert-pivot-confirm').click();
    await page.waitForTimeout(150);

    // Classic layout: header "Sum of Sales" in F1, no per-quarter cols.
    expect(await readCellValue(page, 'E1')).toBe('Region');
    expect(await readCellValue(page, 'F1')).toBe('Sum of Sales');
    expect(await readCellValue(page, 'E2')).toBe('North');
    expect(Number(await readCellValue(page, 'F2'))).toBe(220);
    expect(await readCellValue(page, 'E3')).toBe('South');
    expect(Number(await readCellValue(page, 'F3'))).toBe(175);
    expect(await readCellValue(page, 'E4')).toBe('Grand Total');
    expect(Number(await readCellValue(page, 'F4'))).toBe(395);
    // G1 stays empty — no second value column was written.
    expect(await readCellValue(page, 'G1')).toBeFalsy();
  });
});
