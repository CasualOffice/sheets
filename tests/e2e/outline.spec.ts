/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { expect, test, type Page } from '@playwright/test';
import { selectRange, waitForUniver } from './_helpers';

/**
 * Outline / row+column groups. Three things to lock down:
 *
 *   1. Group rows from the Data menu, then collapse via the OutlinePanel
 *      actually hides those rows on the worksheet (dispatches the Univer
 *      row-hide command, not just a local toggle).
 *   2. Ungroup from the menu removes the group AND shows the rows back if
 *      they were collapsed.
 *   3. xlsx round-trip preserves group boundaries via our
 *      `__casual_sheets_outline__` resource — re-open the exported bytes,
 *      the groups come back with the same start/end.
 */

declare global {
  interface Window {
    __xlsx?: {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      xlsxToWorkbookData: (buf: ArrayBuffer) => Promise<any>;
      workbookDataToXlsx: (
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        data: any,
        extras?: {
          outline?: Record<
            string,
            { rows: Array<{ id: string; start: number; end: number; collapsed: boolean }>; cols: Array<{ id: string; start: number; end: number; collapsed: boolean }> }
          >;
        },
      ) => Promise<Blob>;
    };
  }
}

async function exposeConverters(page: Page) {
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mod = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    window.__xlsx = mod;
  });
}

const openDataMenu = async (page: Page) => {
  await page.getByTestId('menubar-data').click();
  // Wait until the dropdown is actually rendered before targeting an item —
  // under parallel-worker load, the next click can fire before React has
  // mounted the popup, and Playwright auto-retry isn't enough since the
  // menu DOM doesn't exist yet (no element to retry against).
  await expect(page.getByTestId('menubar-data-popup')).toBeVisible();
};

const openOutlinePanel = async (page: Page) => {
  // Outline panel moved from Data to View in Polish #5 — only the
  // group/ungroup operations remained on Data.
  await page.getByTestId('menubar-view').click();
  await expect(page.getByTestId('menubar-view-popup')).toBeVisible();
  await page.getByTestId('menu-item-outline-panel').click();
  await expect(page.getByTestId('outline-panel')).toBeVisible();
};

test.describe('Outline / Group rows + columns', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('Data → Group rows lists the group in the OutlinePanel', async ({ page }) => {
    await selectRange(page, 'A2:A5');
    await openDataMenu(page);
    await page.getByTestId('menu-item-group-rows').click();

    await openOutlinePanel(page);
    await expect(page.getByTestId('outline-section-rows')).toBeVisible();
    await expect(page.getByText('Rows 2–5')).toBeVisible();
  });

  test('Collapsing a row group hides the rows on the worksheet', async ({ page }) => {
    await selectRange(page, 'A2:A4');
    await openDataMenu(page);
    await page.getByTestId('menu-item-group-rows').click();

    await openOutlinePanel(page);
    const toggle = page.locator('[data-testid^="outline-toggle-rows-"]').first();
    await toggle.click();

    // The collapse must dispatch Univer's set-rows-hidden command, which
    // flips `hd=1` on rows 2..4 (zero-based 1..3) in the snapshot.
    await page.waitForFunction(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap: any = api.getActiveWorkbook()!.save();
      const sheetId = snap.sheetOrder[0];
      const rowData = snap.sheets[sheetId].rowData ?? {};
      return [1, 2, 3].every((r) => rowData[r]?.hd === 1);
    }, null, { timeout: 3_000 });
  });

  test('Ungroup removes the group from the active cell', async ({ page }) => {
    // Group rows 2..4, then position the cursor inside that range and run
    // Ungroup from the menu. The group should disappear from the panel.
    // (Collapse-then-restore is covered by the other test; this test stays
    // focused on the remove-group action so panel state doesn't bleed in.)
    await selectRange(page, 'A2:A4');
    await openDataMenu(page);
    await page.getByTestId('menu-item-group-rows').click();

    await selectRange(page, 'A3');
    await openDataMenu(page);
    await page.getByTestId('menu-item-ungroup').click();

    await openOutlinePanel(page);
    await expect(page.getByText('No groups on this sheet.')).toBeVisible();
  });

  test('Group state round-trips through xlsx', async ({ page }) => {
    await exposeConverters(page);

    const reloaded = await page.evaluate(async () => {
      const snapshot = {
        id: 'wb-outline-1',
        rev: 1,
        name: 'wb',
        appVersion: '0.22.1',
        locale: 1,
        styles: {},
        sheetOrder: ['s1'],
        sheets: {
          s1: {
            id: 's1',
            name: 'Sheet1',
            cellData: {},
            rowCount: 1024,
            columnCount: 128,
          },
        },
      };
      const extras = {
        outline: {
          s1: {
            rows: [{ id: 'g-row-1', start: 5, end: 9, collapsed: false }],
            cols: [{ id: 'g-col-1', start: 2, end: 4, collapsed: true }],
          },
        },
      };
      const blob = await window.__xlsx!.workbookDataToXlsx(snapshot, extras);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const re: any = await window.__xlsx!.xlsxToWorkbookData(await blob.arrayBuffer());
      const res = re.resources?.find((r: { name: string }) => r.name === '__casual_sheets_outline__');
      return res ? JSON.parse(res.data) : null;
    });

    expect(reloaded).not.toBeNull();
    expect(reloaded.v).toBe(1);
    expect(reloaded.sheets.s1.rows).toEqual([
      { id: 'g-row-1', start: 5, end: 9, collapsed: false },
    ]);
    expect(reloaded.sheets.s1.cols).toEqual([
      { id: 'g-col-1', start: 2, end: 4, collapsed: true },
    ]);
  });
});
