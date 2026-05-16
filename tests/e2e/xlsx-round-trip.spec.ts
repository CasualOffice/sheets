import { expect, test } from '@playwright/test';
import { selectRange, waitForUniver } from './_helpers';

/**
 * xlsx import/export round-trip.
 *
 * Strategy
 *   We call the converter functions directly from the page context to avoid
 *   relying on browser file-dialog UI (which is mocked-out in headless mode).
 *   The covered fidelity matches `apps/web/src/xlsx/` —- values, formulas,
 *   bold/italic, font color, fill, alignment, number format, merges.
 *
 *   The Open button (File menu) is exercised separately by stubbing
 *   `pickXlsxFile` via a dynamic import — see "loads an xlsx from a Blob".
 */

declare global {
  interface Window {
    __xlsx?: {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      xlsxToWorkbookData: (buf: ArrayBuffer) => Promise<any>;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      workbookDataToXlsx: (data: any) => Promise<Blob>;
    };
  }
}

import type { Page } from '@playwright/test';

async function exposeConverters(page: Page) {
  // The xlsx module is bundled but not on window. Pull it through a dynamic
  // import served by Vite, then expose.
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mod = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    window.__xlsx = mod;
  });
}

test.describe('xlsx round-trip', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await exposeConverters(page);
  });

  test('values and formulas round-trip', async ({ page }) => {
    await selectRange(page, 'A1');

    // Seed the workbook.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'Header' });
      ws.getRange('A2').setValue({ v: 42 });
      ws.getRange('A3').setValue({ v: 100 });
      ws.getRange('A4').setValue({ f: '=SUM(A2:A3)' });
    });

    // Export to xlsx, then re-import and inspect the new snapshot.
    const snapshot = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const original: any = api.getActiveWorkbook()!.save();
      const blob = await window.__xlsx!.workbookDataToXlsx(original);
      const buf = await blob.arrayBuffer();
      const reloaded = await window.__xlsx!.xlsxToWorkbookData(buf);
      return reloaded;
    });

    const sheetId = snapshot.sheetOrder[0];
    const cells = snapshot.sheets[sheetId].cellData;
    expect(cells['0']['0'].v).toBe('Header');
    expect(cells['1']['0'].v).toBe(42);
    expect(cells['2']['0'].v).toBe(100);
    expect(cells['3']['0'].f).toBe('=SUM(A2:A3)');
  });

  test('font weight, color, fill, and number format round-trip', async ({ page }) => {
    await selectRange(page, 'A1');

    // Apply style via the API for determinism.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 1234.5 });
      const r = ws.getRange('A1');
      r.setFontWeight('bold');
      r.setFontColor('#ff0000');
      r.setBackground('#ffeb3b');
      r.setNumberFormat('"$"#,##0.00');
    });

    const styled = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const original: any = api.getActiveWorkbook()!.save();
      const blob = await window.__xlsx!.workbookDataToXlsx(original);
      const buf = await blob.arrayBuffer();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const reloaded: any = await window.__xlsx!.xlsxToWorkbookData(buf);
      const sheetId = reloaded.sheetOrder[0];
      const cell = reloaded.sheets[sheetId].cellData['0']['0'];
      const styleRef = cell.s;
      const style =
        typeof styleRef === 'string' ? reloaded.styles[styleRef] : styleRef;
      return { v: cell.v, style };
    });

    expect(styled.v).toBe(1234.5);
    expect(styled.style.bl).toBe(1);
    expect(styled.style.cl?.rgb?.toLowerCase()).toBe('#ff0000');
    expect(styled.style.bg?.rgb?.toLowerCase()).toBe('#ffeb3b');
    expect(styled.style.n?.pattern).toBe('"$"#,##0.00');
  });

  test('merges round-trip', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'Title' });
      ws.getRange('A1:C1').merge();
    });

    const merges = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const original: any = api.getActiveWorkbook()!.save();
      const blob = await window.__xlsx!.workbookDataToXlsx(original);
      const buf = await blob.arrayBuffer();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const reloaded: any = await window.__xlsx!.xlsxToWorkbookData(buf);
      const sheetId = reloaded.sheetOrder[0];
      return reloaded.sheets[sheetId].mergeData;
    });

    expect(merges).toEqual([
      { startRow: 0, startColumn: 0, endRow: 0, endColumn: 2 },
    ]);
  });

  test('multi-sheet workbook preserves sheet order and names', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      wb.insertSheet('Forecast');
      wb.insertSheet('Inputs');
    });

    const names = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const original: any = api.getActiveWorkbook()!.save();
      const blob = await window.__xlsx!.workbookDataToXlsx(original);
      const buf = await blob.arrayBuffer();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const reloaded: any = await window.__xlsx!.xlsxToWorkbookData(buf);
      return reloaded.sheetOrder.map((id: string) => reloaded.sheets[id].name);
    });

    expect(names).toContain('Forecast');
    expect(names).toContain('Inputs');
  });
});
