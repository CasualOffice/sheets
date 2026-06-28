import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Watch Window persistence — watches round-trip through xlsx via the
 * `__casual_sheets_watches__` resource sidecar (same channel as pivots/charts).
 * Exercises the real export → import → read path in-page; the resource
 * read/write is also unit-tested in watch-resources.unit.test.ts.
 */

test('watches survive an xlsx export → import round-trip', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  const survived = await page.evaluate(async () => {
    const api = window.__univerAPI!;
    const wb = api.getActiveWorkbook()!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const snapshot: any = wb.save();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const xlsx = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wr = await import(/* @vite-ignore */ '/src/shell/watch-resources.ts' as any);
    const sheetId = snapshot.sheetOrder[0];
    const watches = [
      { id: 'w1', sheetId, sheetName: 'Sheet1', row: 0, col: 2 },
      { id: 'w2', sheetId, sheetName: 'Sheet1', row: 4, col: 1 },
    ];
    const blob = await xlsx.workbookDataToXlsx(snapshot, { watches });
    const buf = await blob.arrayBuffer();
    const reloaded = await xlsx.xlsxToWorkbookData(buf);
    return wr.readWatchesFromSnapshot(reloaded);
  });

  expect(survived).toHaveLength(2);
  expect(survived[0]).toMatchObject({ id: 'w1', row: 0, col: 2 });
  expect(survived[1]).toMatchObject({ id: 'w2', row: 4, col: 1 });
});
