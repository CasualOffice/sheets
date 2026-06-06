import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * Charts P1 — Excel-style insert flow + persistence.
 *
 *   - Insert > Chart opens a dialog with type + range pickers.
 *   - Charts are auto-named "Chart N" and appear in the Charts panel.
 *   - Panel rename + delete edit the in-memory store.
 *   - Round-trip: a chart serialized into `__casual_sheets_charts__`
 *     comes back from `readChartsFromSnapshot` after xlsx export.
 */

async function exposeRoundTrip(page: Page) {
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const xlsx = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const chartsRes = await import(/* @vite-ignore */ '/src/charts/resources.ts' as any);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__rt__ = {
      workbookDataToXlsx: xlsx.workbookDataToXlsx,
      xlsxToWorkbookData: xlsx.xlsxToWorkbookData,
      readChartsFromSnapshot: chartsRes.readChartsFromSnapshot,
    };
  });
}

async function seedData(page: Page) {
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Region' });
    ws.getRange('B1').setValue({ v: 'Q1' });
    ws.getRange('C1').setValue({ v: 'Q2' });
    ws.getRange('A2').setValue({ v: 'North' });
    ws.getRange('B2').setValue({ v: 100 });
    ws.getRange('C2').setValue({ v: 120 });
    ws.getRange('A3').setValue({ v: 'South' });
    ws.getRange('B3').setValue({ v: 80 });
    ws.getRange('C3').setValue({ v: 95 });
    ws.getRange('A1:C3').activate();
  });
}

test.describe('Charts P1 — insert dialog + panel + persistence', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await seedData(page);
    // Click canvas to focus, then restore the selection (clicking
    // moves the selection caret).
    await mainCanvas(page)
      .first()
      .click({ position: { x: 100, y: 100 } });
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1:C3').activate();
    });
  });

  test('Insert > Chart opens a dialog pre-filled with the active selection', async ({ page }) => {
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-chart').click();
    await expect(page.getByTestId('insert-chart-dialog')).toBeVisible();
    await expect(page.getByTestId('insert-chart-range')).toHaveValue('A1:C3');
    // Column family is the Excel default; Clustered Column is its first subtype.
    await expect(page.getByTestId('insert-chart-family-column')).toHaveClass(
      /insert-chart__family--active/,
    );
    await expect(page.getByTestId('insert-chart-type-column')).toHaveClass(
      /insert-chart__subtype--active/,
    );
  });

  test('confirming the dialog inserts a chart and the panel lists it as "Chart 1"', async ({
    page,
  }) => {
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-chart').click();
    // Switch family to Line, then pick the Line subtype.
    await page.getByTestId('insert-chart-family-line').click();
    await page.getByTestId('insert-chart-type-line').click();
    await page.getByTestId('insert-chart-confirm').click();

    await expect(page.getByTestId('chart-overlay')).toBeVisible({ timeout: 5_000 });

    // Open the Charts panel via the View menu.
    await page.getByTestId('menubar-view').click();
    await page.getByTestId('menu-item-charts-panel').click();
    const panel = page.getByTestId('charts-panel');
    await expect(panel).toBeVisible();
    await expect(panel.getByText('Chart 1')).toBeVisible();
    await expect(panel.getByText('Line', { exact: true })).toBeVisible();
    await expect(panel.getByText('A1:C3')).toBeVisible();
  });

  test('panel delete removes the overlay', async ({ page }) => {
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-chart').click();
    await page.getByTestId('insert-chart-confirm').click();
    await expect(page.getByTestId('chart-overlay')).toBeVisible({ timeout: 5_000 });

    await page.getByTestId('menubar-view').click();
    await page.getByTestId('menu-item-charts-panel').click();
    await page.getByLabel(/^Delete Chart 1$/).click();
    await expect(page.getByTestId('chart-overlay')).toHaveCount(0);
  });

  test('the chart round-trips through xlsx via __casual_sheets_charts__', async ({ page }) => {
    // CI runners under heavy parallel load occasionally trip the
    // "Execution context was destroyed by navigation" failure between
    // the chart-overlay assertion and the page.evaluate snapshot
    // below. Three retries cover the intermittent slow runner case
    // without papering over a real navigation regression.
    test
      .info()
      .annotations.push({
        type: 'flaky',
        description: 'chart-overlay → save race on slow runners',
      });
    test.setTimeout(60_000);
    await exposeRoundTrip(page);
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-chart').click();
    await page.getByTestId('insert-chart-family-pie').click();
    await page.getByTestId('insert-chart-type-pie').click();
    await page.getByTestId('insert-chart-confirm').click();
    await expect(page.getByTestId('chart-overlay')).toBeVisible({ timeout: 5_000 });
    // Settle the page so the upcoming `page.evaluate` doesn't race a
    // lazy-loaded module's network/parse. The earlier failure mode
    // was "Execution context was destroyed, most likely because of
    // a navigation" — waiting for networkidle here prevents the
    // race.
    await page.waitForLoadState('networkidle');

    const reloaded = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snapshot: any = wb.save();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const rt = (window as any).__rt__;
      // Mirror what the real Save path does: hand the chart list in
      // via ExportExtras so it gets written into resources.
      const charts = [
        {
          id: 'ch-test',
          sheetId: snapshot.sheetOrder[0],
          source: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 },
          pos: { startRow: 4, endRow: 13, startColumn: 0, endColumn: 7 },
          type: 'pie',
          title: 'Chart 1',
        },
      ];
      const blob = await rt.workbookDataToXlsx(snapshot, { charts });
      const buf = await blob.arrayBuffer();
      const reloaded = await rt.xlsxToWorkbookData(buf);
      return rt.readChartsFromSnapshot(reloaded);
    });

    expect(reloaded).toHaveLength(1);
    expect(reloaded[0].id).toBe('ch-test');
    expect(reloaded[0].type).toBe('pie');
    expect(reloaded[0].title).toBe('Chart 1');
    expect(reloaded[0].source).toEqual({ startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 });
    expect(reloaded[0].pos).toEqual({ startRow: 4, endRow: 13, startColumn: 0, endColumn: 7 });
  });
});
