import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Native pivot export (behind the `cs-native-pivots` opt-in). With the flag on,
 * exporting a workbook with a simple in-app pivot emits real xl/pivotTables
 * parts; with it off, export is unchanged (flat cells only). The OOXML
 * generation is unit-tested in pivot-export.unit.test.ts; this checks the live
 * export wiring + the flag gate.
 */

async function runExport(page: import('@playwright/test').Page, flagOn: boolean) {
  return page.evaluate(async (enable) => {
    if (enable) localStorage.setItem('cs-native-pivots', '1');
    else localStorage.removeItem('cs-native-pivots');
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Region' });
    ws.getRange('B1').setValue({ v: 'Quarter' });
    ws.getRange('C1').setValue({ v: 'Sales' });
    const rows = [
      ['North', 'Q1', 100],
      ['South', 'Q1', 80],
      ['North', 'Q2', 120],
      ['South', 'Q2', 95],
    ];
    rows.forEach((r, i) => {
      ws.getRange('A' + (i + 2)).setValue({ v: r[0] });
      ws.getRange('B' + (i + 2)).setValue({ v: r[1] });
      ws.getRange('C' + (i + 2)).setValue({ v: r[2] });
    });
    const wb = api.getActiveWorkbook()!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const snapshot: any = wb.save();
    const sheetId = snapshot.sheetOrder[0];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const xlsx = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    const model = {
      id: 'pt1',
      sourceSheetId: sheetId,
      source: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 },
      targetSheetId: sheetId,
      target: { row: 0, column: 4 },
      rows: [{ column: 0 }],
      cols: [],
      values: [{ column: 2, agg: 'sum' }],
    };
    const blob = await xlsx.workbookDataToXlsx(snapshot, { pivots: [model] });
    // Zip entry filenames are stored uncompressed in the local file headers,
    // so a byte-scan reliably detects whether the native pivot part exists
    // (avoids resolving 'jszip' as a bare specifier in the browser).
    const bytes = new Uint8Array(await blob.arrayBuffer());
    let raw = '';
    for (let i = 0; i < bytes.length; i++) raw += String.fromCharCode(bytes[i]);
    return { hasPivotTable: raw.includes('pivotTables/pivotTable1.xml') };
  }, flagOn);
}

test('flag on → native xl/pivotTables emitted for a simple pivot', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  const out = await runExport(page, true);
  expect(out.hasPivotTable).toBe(true);
});

test('flag off → no native pivot parts (default, unchanged export)', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  const out = await runExport(page, false);
  expect(out.hasPivotTable).toBe(false);
});
