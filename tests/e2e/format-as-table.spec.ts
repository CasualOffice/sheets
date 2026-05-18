import { expect, test } from '@playwright/test';
import { selectRange, waitForUniver } from './_helpers';

declare global {
  interface Window {
    __getTableStyleId__?: (tableId: string) => string | undefined;
  }
}

test.describe('Format as Table', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    // Seed a tabular range A1:C4 with a header row.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'Name' });
      ws.getRange('B1').setValue({ v: 'Score' });
      ws.getRange('C1').setValue({ v: 'Rank' });
      ws.getRange('A2').setValue({ v: 'Ada' });
      ws.getRange('B2').setValue({ v: 92 });
      ws.getRange('C2').setValue({ v: 1 });
      ws.getRange('A3').setValue({ v: 'Bob' });
      ws.getRange('B3').setValue({ v: 88 });
      ws.getRange('C3').setValue({ v: 2 });
      ws.getRange('A4').setValue({ v: 'Cleo' });
      ws.getRange('B4').setValue({ v: 75 });
      ws.getRange('C4').setValue({ v: 3 });
    });
  });

  test('Format as Table — default click applies the first theme over A1:C4', async ({ page }) => {
    await selectRange(page, 'A1:C4');
    await page.getByTestId('ribbon-dropdown-format-as-table-apply').click();
    await page.waitForTimeout(500);

    const tables = await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const all = (wb as any).getTableList?.() ?? [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return all.map((t: any) => ({
        styleId: window.__getTableStyleId__?.(t.id),
        range: t.range,
      }));
    });
    expect(tables.length).toBe(1);
    expect(tables[0].styleId).toBe('table-default-0');
    expect(tables[0].range.startRow).toBe(0);
    expect(tables[0].range.endRow).toBe(3);
  });

  test('Theme picker — choosing Green applies table-default-2', async ({ page }) => {
    await selectRange(page, 'A1:C4');
    await page.getByTestId('ribbon-dropdown-format-as-table-caret').click();
    await page
      .getByTestId('ribbon-dropdown-format-as-table-item-table-default-2')
      .click();
    await page.waitForTimeout(500);

    const styleId = await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const all = (wb as any).getTableList?.() ?? [];
      return all[0] ? window.__getTableStyleId__?.(all[0].id) : undefined;
    });
    expect(styleId).toBe('table-default-2');
  });

  test('Single-cell selection inside the data block auto-detects the bounds', async ({ page }) => {
    // Click into B2 only — the action should expand to A1:C4.
    await selectRange(page, 'B2');
    await page.getByTestId('ribbon-dropdown-format-as-table-apply').click();
    await page.waitForTimeout(500);

    const range = await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const all = (wb as any).getTableList?.() ?? [];
      return all[0]?.range;
    });
    expect(range).toBeTruthy();
    expect(range.startRow).toBe(0);
    expect(range.startColumn).toBe(0);
    expect(range.endRow).toBe(3);
    expect(range.endColumn).toBe(2);
  });

  test('Rapid double-click on Format as Table creates ONE table, not two', async ({ page }) => {
    // Repro for the "duplicate table created in background thread" bug —
    // a fast user-click sequence used to race the lazy-plugin load and
    // dispatch `add-table` twice. The in-flight guard in formatAsTable
    // (`apps/web/src/shell/tab-actions.ts`) plus the awaited plugin
    // ensure should now coalesce them.
    await selectRange(page, 'A1:C4');
    const btn = page.getByTestId('ribbon-dropdown-format-as-table-apply');
    // Two clicks within the same event-loop tick — the second click
    // arrives while the first is still awaiting ensurePlugin/addTable.
    await Promise.all([btn.click(), btn.click()]);
    // Give both clicks time to complete their async work before
    // counting tables.
    await page.waitForTimeout(800);

    const count = await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return ((wb as any).getTableList?.() ?? []).length;
    });
    expect(count).toBe(1);
  });
});
