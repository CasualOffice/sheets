import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

declare global {
  interface Window {
    __odsFreeze?: typeof import('../../apps/web/src/ods');
  }
}

test.describe('ods freeze panes round-trip', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const mod = await import(/* @vite-ignore */ '/src/ods/index.ts' as any);
      window.__odsFreeze = mod;
    });
  });

  test('snapshot freeze metadata survives ods export and import', async ({ page }) => {
    const result = await page.evaluate(async () => {
      const snapshot = {
        id: 'wb-freeze-1',
        rev: 1,
        name: 'wb',
        appVersion: '0.22.1',
        locale: 1,
        styles: {},
        sheetOrder: ['s1'],
        sheets: {
          s1: {
            id: 's1',
            name: 'Data',
            cellData: { 0: { 0: { v: 'x' } } },
            rowCount: 100,
            columnCount: 26,
            freeze: { xSplit: 1, ySplit: 2, startRow: 2, startColumn: 1 },
          },
        },
      };

      const blob = await window.__odsFreeze!.workbookDataToOds(snapshot as any);
      const buf = await blob.arrayBuffer();
      const reloaded = await window.__odsFreeze!.odsToWorkbookData(buf);
      const sheetId = reloaded.sheetOrder[0];
      return reloaded.sheets[sheetId]?.freeze ?? null;
    });

    expect(result).toEqual({ xSplit: 1, ySplit: 2, startRow: 2, startColumn: 1 });
  });
});
