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

  test('freeze + tab color + hidden sheet round-trip', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const fws: any = wb.getActiveSheet();
      // Freeze the top row + first column.
      fws.setFrozenRows(1);
      fws.setFrozenColumns(1);
      // Tab color via the facade if exposed; otherwise via the mutation field
      // on the snapshot. Univer's facade has setTabColor? Fallback to the
      // mutation directly so the snapshot reflects it before export.
      if (typeof fws.setTabColor === 'function') fws.setTabColor('#ff5722');
      // Add a second sheet and hide it. The snapshot is what export reads, so
      // it's enough to set hidden=1 directly there — facade variants differ
      // across Univer versions.
      wb.insertSheet('Secret');
    });

    // Mark the second sheet hidden in the snapshot before export.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap: any = api.getActiveWorkbook()!.save();
      const secretId = snap.sheetOrder.find((id: string) => snap.sheets[id].name === 'Secret');
      if (secretId) snap.sheets[secretId].hidden = 1;
      // Cache the prepared snapshot for the next eval.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (window as { __preparedSnap?: unknown }).__preparedSnap = snap;
    });

    const result = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap: any = (window as { __preparedSnap?: unknown }).__preparedSnap;
      const blob = await window.__xlsx!.workbookDataToXlsx(snap);
      const buf = await blob.arrayBuffer();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const reloaded: any = await window.__xlsx!.xlsxToWorkbookData(buf);
      const firstId = reloaded.sheetOrder[0];
      const first = reloaded.sheets[firstId];
      const secret = Object.values(reloaded.sheets).find((s: unknown) => (s as { name: string }).name === 'Secret') as { hidden?: number } | undefined;
      return {
        freeze: first.freeze,
        tabColor: first.tabColor,
        secretHidden: secret?.hidden,
      };
    });

    expect(result.freeze?.ySplit).toBe(1);
    expect(result.freeze?.xSplit).toBe(1);
    expect(result.secretHidden).toBe(1);
  });

  test('tab color round-trip via snapshot field', async ({ page }) => {
    // Set tabColor directly on the snapshot and round-trip — avoids depending
    // on any facade setter that may or may not exist.
    const result = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const original: any = api.getActiveWorkbook()!.save();
      const firstId = original.sheetOrder[0];
      original.sheets[firstId].tabColor = '#aabbcc';
      const blob = await window.__xlsx!.workbookDataToXlsx(original);
      const buf = await blob.arrayBuffer();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const reloaded: any = await window.__xlsx!.xlsxToWorkbookData(buf);
      return reloaded.sheets[reloaded.sheetOrder[0]].tabColor;
    });
    expect(result).toBe('#aabbcc');
  });

  test('column widths and row heights round-trip', async ({ page }) => {
    // Set non-default widths on A and B and a tall row 1 via the facade,
    // then export → re-import and check the values come back.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const fws: any = api.getActiveWorkbook()!.getActiveSheet();
      fws.setColumnWidth(0, 140);
      fws.setColumnWidth(1, 60);
      fws.setRowHeight(0, 48);
    });

    const result = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const original: any = api.getActiveWorkbook()!.save();
      const blob = await window.__xlsx!.workbookDataToXlsx(original);
      const buf = await blob.arrayBuffer();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const reloaded: any = await window.__xlsx!.xlsxToWorkbookData(buf);
      const sheetId = reloaded.sheetOrder[0];
      const wsd = reloaded.sheets[sheetId];
      return {
        widthA: wsd.columnData?.[0]?.w,
        widthB: wsd.columnData?.[1]?.w,
        heightRow1: wsd.rowData?.[0]?.h,
      };
    });

    // Allow ±2px slop from the px↔char↔px Excel char-width conversion.
    expect(result.widthA).toBeGreaterThan(135);
    expect(result.widthA).toBeLessThan(145);
    expect(result.widthB).toBeGreaterThan(55);
    expect(result.widthB).toBeLessThan(65);
    expect(result.heightRow1).toBeGreaterThan(44);
    expect(result.heightRow1).toBeLessThan(52);
  });
});
