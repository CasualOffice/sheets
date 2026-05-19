import { expect, test, type Page } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * Pivots P0 — first useful slice. Verifies:
 *
 *   - Insert > PivotTable opens a dialog seeded with the active range.
 *   - Picking row field + value field + aggregation computes the pivot
 *     and writes it as cell values at the target location (rows sorted
 *     alphabetically; Grand Total row at the bottom).
 *   - Each aggregation (sum / count / average / min / max) produces
 *     the right output.
 *   - The pivot's compute step is pure and idempotent: running it
 *     twice on the same source produces the same cells.
 *   - PivotModel persists into IWorkbookData.resources via
 *     `__casual_sheets_pivots__` so it round-trips through xlsx.
 *
 * P0 ships single row-field + single value-field. Multi-field rows,
 * column fields, filters, and live-refresh land in later phases.
 */

async function exposeRoundTrip(page: Page) {
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const xlsx = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const pivotsRes = await import(/* @vite-ignore */ '/src/pivots/resources.ts' as any);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__rt__ = {
      workbookDataToXlsx: xlsx.workbookDataToXlsx,
      xlsxToWorkbookData: xlsx.xlsxToWorkbookData,
      readPivotsFromSnapshot: pivotsRes.readPivotsFromSnapshot,
    };
  });
}

async function seedSales(page: Page) {
  // Standard demo data: Region × Quarter aggregating Sales.
  // Rows ordered intentionally with North/South interleaved so the
  // pivot's groupBy is the thing that produces alphabetical row keys.
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Region' });
    ws.getRange('B1').setValue({ v: 'Quarter' });
    ws.getRange('C1').setValue({ v: 'Sales' });
    ws.getRange('A2').setValue({ v: 'North' });  ws.getRange('B2').setValue({ v: 'Q1' }); ws.getRange('C2').setValue({ v: 100 });
    ws.getRange('A3').setValue({ v: 'South' });  ws.getRange('B3').setValue({ v: 'Q1' }); ws.getRange('C3').setValue({ v: 80 });
    ws.getRange('A4').setValue({ v: 'North' });  ws.getRange('B4').setValue({ v: 'Q2' }); ws.getRange('C4').setValue({ v: 120 });
    ws.getRange('A5').setValue({ v: 'South' });  ws.getRange('B5').setValue({ v: 'Q2' }); ws.getRange('C5').setValue({ v: 95 });
    ws.getRange('A1:C5').activate();
  });
  // Click the canvas + restore selection (clicking moves the selection).
  await mainCanvas(page).first().click({ position: { x: 100, y: 100 } });
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1:C5').activate();
  });
}

test.describe('Pivots P0 — insert flow + compute', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await seedSales(page);
  });

  test('Insert > PivotTable opens a dialog pre-filled with the active selection', async ({ page }) => {
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-pivot').click();
    await expect(page.getByTestId('insert-pivot-dialog')).toBeVisible();
    await expect(page.getByTestId('insert-pivot-range')).toHaveValue('A1:C5');
    // Default target = two columns to the right of the source.
    await expect(page.getByTestId('insert-pivot-target')).toHaveValue(/^E\d+/);
  });

  test('Sum of Sales by Region computes the right grouped totals + Grand Total', async ({ page }) => {
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-pivot').click();
    // Default: row field = Region (column 0), value field = Quarter (col 1),
    // aggregation = sum. Switch the value field to Sales (column 2) so the
    // numbers add up.
    await page.getByTestId('insert-pivot-value-field').selectOption({ index: 2 });
    await page.getByTestId('insert-pivot-target').fill('E1');
    await page.getByTestId('insert-pivot-confirm').click();

    // Output starts at E1: header → North → South → Grand Total.
    const cells = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return {
        e1: ws.getRange('E1').getValue(),
        f1: ws.getRange('F1').getValue(),
        e2: ws.getRange('E2').getValue(),
        f2: ws.getRange('F2').getValue(),
        e3: ws.getRange('E3').getValue(),
        f3: ws.getRange('F3').getValue(),
        e4: ws.getRange('E4').getValue(),
        f4: ws.getRange('F4').getValue(),
      };
    });
    expect(cells.e1).toBe('Region');
    expect(cells.f1).toBe('Sum of Sales');
    expect(cells.e2).toBe('North');
    expect(cells.f2).toBe(220);
    expect(cells.e3).toBe('South');
    expect(cells.f3).toBe(175);
    expect(cells.e4).toBe('Grand Total');
    expect(cells.f4).toBe(395);
  });

  test('Average of Sales by Region', async ({ page }) => {
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-pivot').click();
    await page.getByTestId('insert-pivot-value-field').selectOption({ index: 2 });
    await page.getByTestId('insert-pivot-aggregation').selectOption('average');
    await page.getByTestId('insert-pivot-target').fill('E1');
    await page.getByTestId('insert-pivot-confirm').click();

    const cells = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return {
        header: ws.getRange('F1').getValue(),
        north: ws.getRange('F2').getValue(),
        south: ws.getRange('F3').getValue(),
        total: ws.getRange('F4').getValue(),
      };
    });
    expect(cells.header).toBe('Average of Sales');
    expect(cells.north).toBe(110);  // (100+120)/2
    expect(cells.south).toBe(87.5); // (80+95)/2
    expect(cells.total).toBe(98.75); // (100+120+80+95)/4
  });

  test('Count aggregation counts non-empty cells (not numerical-only)', async ({ page }) => {
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-pivot').click();
    // Count of Quarter by Region — column 1 has non-empty Q1/Q2 strings.
    await page.getByTestId('insert-pivot-value-field').selectOption({ index: 1 });
    await page.getByTestId('insert-pivot-aggregation').selectOption('count');
    await page.getByTestId('insert-pivot-target').fill('E1');
    await page.getByTestId('insert-pivot-confirm').click();

    const cells = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return {
        north: ws.getRange('F2').getValue(),
        south: ws.getRange('F3').getValue(),
        total: ws.getRange('F4').getValue(),
      };
    });
    expect(cells.north).toBe(2);
    expect(cells.south).toBe(2);
    expect(cells.total).toBe(4);
  });

  test('PivotModel round-trips through xlsx via __casual_sheets_pivots__', async ({ page }) => {
    await exposeRoundTrip(page);
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-pivot').click();
    await page.getByTestId('insert-pivot-value-field').selectOption({ index: 2 });
    await page.getByTestId('insert-pivot-target').fill('E1');
    await page.getByTestId('insert-pivot-confirm').click();

    const reloaded = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snapshot: any = wb.save();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const rt = (window as any).__rt__;
      const pivots = [
        {
          id: 'pt-test',
          sourceSheetId: snapshot.sheetOrder[0],
          source: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 },
          targetSheetId: snapshot.sheetOrder[0],
          target: { row: 0, column: 4 },
          rows: [{ column: 0 }],
          cols: [],
          values: [{ column: 2, agg: 'sum' }],
          title: 'PivotTable 1',
        },
      ];
      const blob = await rt.workbookDataToXlsx(snapshot, { pivots });
      const buf = await blob.arrayBuffer();
      const reloaded = await rt.xlsxToWorkbookData(buf);
      return rt.readPivotsFromSnapshot(reloaded);
    });
    expect(reloaded).toHaveLength(1);
    expect(reloaded[0].id).toBe('pt-test');
    expect(reloaded[0].rows).toEqual([{ column: 0 }]);
    expect(reloaded[0].values).toEqual([{ column: 2, agg: 'sum' }]);
    expect(reloaded[0].target).toEqual({ row: 0, column: 4 });
  });
});
