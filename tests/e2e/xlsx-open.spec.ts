import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

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

/**
 * Regression guard for the React-root-unmount race that left the grid blank
 * after File → Open. See commit 45eab39: Univer is created once and the
 * workbook unit is swapped on snapshot change rather than rebuilding Univer.
 */
test('File → Open replaces grid contents with the picked file', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(`pageerror: ${e.message}`));
  page.on('console', (m) => {
    if (m.type() === 'error' || (m.type() === 'warning' && m.text().includes('synchronously unmount'))) {
      errors.push(`[${m.type()}] ${m.text()}`);
    }
  });

  await page.goto('/');
  await waitForUniver(page);

  // Hoist the bundled xlsx converters onto window so we can fabricate a fixture.
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const xlsx = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    window.__xlsx = xlsx;
  });

  // Build a fixture .xlsx whose values are unique strings so we can detect the swap.
  const fixtureBytes: number[] = await page.evaluate(async () => {
    const data = {
      id: 'imported-1',
      rev: 1,
      name: 'imported',
      appVersion: '0.22.1',
      locale: 1,
      styles: {},
      sheetOrder: ['imp-1'],
      sheets: {
        'imp-1': {
          id: 'imp-1',
          name: 'Imported',
          cellData: { 0: { 0: { v: 'FILE_X' }, 1: { v: 'FILE_Y' } } },
          rowCount: 1024,
          columnCount: 128,
        },
      },
    };
    const blob = await window.__xlsx!.workbookDataToXlsx(data);
    return Array.from(new Uint8Array(await blob.arrayBuffer()));
  });
  const fs = await import('node:fs');
  const fixture = '/tmp/casual-sheets-open-test.xlsx';
  fs.writeFileSync(fixture, Buffer.from(fixtureBytes));

  // Type into the blank workbook so we can distinguish "pre-open" vs "post-open".
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'TYPED_A' });
    ws.getRange('B1').setValue({ v: 'TYPED_B' });
  });

  // File → Open and inject the fixture.
  await page.getByTestId('menubar-file').click();
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.getByTestId('menu-item-open').click(),
  ]);
  await chooser.setFiles(fixture);

  // The xlsx parser runs in a Web Worker (slower on CI runners). Wait
  // until the new workbook is actually mounted instead of a fixed
  // sleep — otherwise the assertion races the post-overlay swap.
  await page.waitForFunction(
    () => {
      const api = window.__univerAPI;
      if (!api) return false;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()?.getActiveSheet();
      return ws?.getRange('A1').getValue() === 'FILE_X';
    },
    null,
    { timeout: 15_000 },
  );

  const after = await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return {
      a1: ws.getRange('A1').getValue(),
      b1: ws.getRange('B1').getValue(),
    };
  });

  expect(after.a1).toBe('FILE_X');
  expect(after.b1).toBe('FILE_Y');
  expect(errors, errors.join('\n')).toEqual([]);
});
