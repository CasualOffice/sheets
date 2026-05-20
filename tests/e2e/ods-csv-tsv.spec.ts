import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

declare global {
  interface Window {
    __odsMod?: typeof import('../../apps/web/src/ods');
  }
}

test.describe('Spreadsheet formats — ods / csv / tsv round-trip', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const mod = await import(/* @vite-ignore */ '/src/ods/index.ts' as any);
      window.__odsMod = mod;
    });
  });

  test('ods values + merges round-trip', async ({ page }) => {
    const result = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'Header' });
      ws.getRange('B1').setValue({ v: 42 });
      ws.getRange('A2').setValue({ v: 'Row 2' });
      ws.getRange('A2:B2').merge();

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap: any = api.getActiveWorkbook()!.save();
      const blob = await window.__odsMod!.workbookDataToOds(snap);
      const buf = await blob.arrayBuffer();
      const reloaded = await window.__odsMod!.odsToWorkbookData(buf);
      const id = reloaded.sheetOrder[0];
      const wsd = reloaded.sheets[id];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const cells = wsd!.cellData as any;
      return {
        a1: cells['0']['0'].v,
        b1: cells['0']['1'].v,
        a2: cells['1']['0'].v,
        merges: wsd!.mergeData,
      };
    });
    expect(result.a1).toBe('Header');
    expect(result.b1).toBe(42);
    expect(result.a2).toBe('Row 2');
    expect(result.merges).toEqual([
      { startRow: 1, startColumn: 0, endRow: 1, endColumn: 1 },
    ]);
  });

  test('ods number formats round-trip through snapshot styles', async ({ page }) => {
    const result = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb: any = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = wb.getActiveSheet();
      ws.getRange('A1').setValue({ v: 1234.56, s: { n: { pattern: '"$"#,##0.00' } } });
      ws.getRange('A2').setValue({ v: 0.25, s: { n: { pattern: '0.00%' } } });

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap: any = wb.save();
      const blob = await window.__odsMod!.workbookDataToOds(snap);
      const buf = await blob.arrayBuffer();
      const reloaded = await window.__odsMod!.odsToWorkbookData(buf);
      const id = reloaded.sheetOrder[0];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const cells = reloaded.sheets[id]!.cellData as any;
      const a1 = cells['0']['0'];
      const a2 = cells['1']['0'];
      return {
        a1Pattern: reloaded.styles?.[a1.s]?.n?.pattern,
        a2Pattern: reloaded.styles?.[a2.s]?.n?.pattern,
        a1Value: a1.v,
        a2Value: a2.v,
      };
    });

    expect(result.a1Pattern).toBe('"$"#,##0.00');
    expect(result.a2Pattern).toBe('0.00%');
    expect(result.a1Value).toBe(1234.56);
    expect(result.a2Value).toBe(0.25);
  });

  test('csv export matches expected text and re-imports', async ({ page }) => {
    const result = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'name' });
      ws.getRange('B1').setValue({ v: 'qty' });
      ws.getRange('A2').setValue({ v: 'apple, gala' });   // forces quoting
      ws.getRange('B2').setValue({ v: 12 });

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap: any = api.getActiveWorkbook()!.save();
      const blob = await window.__odsMod!.workbookDataToDelimited(snap, 'csv');
      const text = await blob.text();

      // Re-import via csvToWorkbookData.
      const buf = new TextEncoder().encode(text).buffer;
      const reloaded = await window.__odsMod!.csvToWorkbookData(buf);
      const id = reloaded.sheetOrder[0];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const cells = reloaded.sheets[id]!.cellData as any;
      return {
        text,
        a1: cells['0']['0'].v,
        a2: cells['1']['0'].v,
        b2: cells['1']['1'].v,
      };
    });
    expect(result.text).toContain('"apple, gala"');
    expect(result.a1).toBe('name');
    expect(result.a2).toBe('apple, gala');
    expect(result.b2).toBe(12);
  });

  test('tsv export uses tab separator and re-imports', async ({ page }) => {
    const result = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'a' });
      ws.getRange('B1').setValue({ v: 'b,c' });            // comma fine in TSV
      ws.getRange('A2').setValue({ v: 1 });
      ws.getRange('B2').setValue({ v: 2 });

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap: any = api.getActiveWorkbook()!.save();
      const blob = await window.__odsMod!.workbookDataToDelimited(snap, 'tsv');
      const text = await blob.text();

      const buf = new TextEncoder().encode(text).buffer;
      const reloaded = await window.__odsMod!.tsvToWorkbookData(buf);
      const id = reloaded.sheetOrder[0];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const cells = reloaded.sheets[id]!.cellData as any;
      return {
        text,
        b1: cells['0']['1'].v,
        a2: cells['1']['0'].v,
      };
    });
    expect(result.text).toContain('\t');
    // Comma in TSV doesn't need quoting.
    expect(result.text).toContain('b,c');
    expect(result.b1).toBe('b,c');
    expect(result.a2).toBe(1);
  });
});
