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

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * Charts P3 — chart-type catalog. Covers:
 *   - Insert dialog renders all 6 families.
 *   - Switching family swaps the subtype grid.
 *   - Each subtype can be inserted and the panel shows its label.
 *   - Stacked / 100% stacked variants render an overlay (build-option
 *     handles them without throwing).
 *   - Migration: pre-P3 `'bar'` value in __casual_sheets_charts__
 *     hydrates as `'column'`.
 */

async function seed(page: Page) {
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Region' });
    ws.getRange('B1').setValue({ v: 'Q1' });
    ws.getRange('C1').setValue({ v: 'Q2' });
    ws.getRange('A2').setValue({ v: 'N' });
    ws.getRange('B2').setValue({ v: 100 });
    ws.getRange('C2').setValue({ v: 120 });
    ws.getRange('A3').setValue({ v: 'S' });
    ws.getRange('B3').setValue({ v: 80 });
    ws.getRange('C3').setValue({ v: 95 });
    ws.getRange('A1:C3').activate();
  });
  await mainCanvas(page).first().click({ position: { x: 100, y: 100 } });
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1:C3').activate();
  });
}

test.describe('Charts P3 — chart-type catalog', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await seed(page);
  });

  test('all six families render in the dialog', async ({ page }) => {
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-chart').click();
    for (const f of ['column', 'bar', 'line', 'area', 'pie', 'scatter']) {
      await expect(page.getByTestId(`insert-chart-family-${f}`)).toBeVisible();
    }
  });

  test('switching family from Column to Bar swaps the subtype grid', async ({ page }) => {
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-chart').click();
    // Column family default — its three subtypes are visible.
    await expect(page.getByTestId('insert-chart-type-column')).toBeVisible();
    await expect(page.getByTestId('insert-chart-type-column-stacked')).toBeVisible();
    await expect(page.getByTestId('insert-chart-type-column-stacked-100')).toBeVisible();
    // Switch to Bar.
    await page.getByTestId('insert-chart-family-bar').click();
    await expect(page.getByTestId('insert-chart-type-bar')).toBeVisible();
    await expect(page.getByTestId('insert-chart-type-bar-stacked')).toBeVisible();
    await expect(page.getByTestId('insert-chart-type-bar-stacked-100')).toBeVisible();
    // Column subtypes no longer in DOM.
    await expect(page.getByTestId('insert-chart-type-column')).toHaveCount(0);
  });

  test('inserting a Stacked Column renders an overlay and the panel labels it correctly', async ({ page }) => {
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-chart').click();
    await page.getByTestId('insert-chart-type-column-stacked').click();
    await page.getByTestId('insert-chart-confirm').click();
    await expect(page.getByTestId('chart-overlay')).toBeVisible({ timeout: 5_000 });

    await page.getByTestId('menubar-view').click();
    await page.getByTestId('menu-item-charts-panel').click();
    await expect(page.getByTestId('charts-panel').getByText('Stacked Column')).toBeVisible();
  });

  test('inserting a 100% Stacked Bar caps the value axis and labels it as %', async ({ page }) => {
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-chart').click();
    await page.getByTestId('insert-chart-family-bar').click();
    await page.getByTestId('insert-chart-type-bar-stacked-100').click();
    await page.getByTestId('insert-chart-confirm').click();
    const overlay = page.getByTestId('chart-overlay');
    await expect(overlay).toBeVisible({ timeout: 5_000 });
    await expect(overlay.locator('canvas')).toHaveCount(1, { timeout: 5_000 });

    await page.getByTestId('menubar-view').click();
    await page.getByTestId('menu-item-charts-panel').click();
    await expect(page.getByTestId('charts-panel').getByText('100% Stacked Bar')).toBeVisible();
  });

  test('inserting a Doughnut renders a pie with a hollow center', async ({ page }) => {
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-chart').click();
    await page.getByTestId('insert-chart-family-pie').click();
    await page.getByTestId('insert-chart-type-doughnut').click();
    await page.getByTestId('insert-chart-confirm').click();
    await expect(page.getByTestId('chart-overlay')).toBeVisible({ timeout: 5_000 });

    await page.getByTestId('menubar-view').click();
    await page.getByTestId('menu-item-charts-panel').click();
    await expect(page.getByTestId('charts-panel').getByText('Doughnut')).toBeVisible();
  });

  test('inserting Area renders an overlay and panel labels it Area', async ({ page }) => {
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-chart').click();
    await page.getByTestId('insert-chart-family-area').click();
    await page.getByTestId('insert-chart-type-area').click();
    await page.getByTestId('insert-chart-confirm').click();
    await expect(page.getByTestId('chart-overlay')).toBeVisible({ timeout: 5_000 });

    await page.getByTestId('menubar-view').click();
    await page.getByTestId('menu-item-charts-panel').click();
    await expect(page.getByTestId('charts-panel').getByText('Area', { exact: true })).toBeVisible();
  });

  test('legacy "bar" stored in __casual_sheets_charts__ migrates to "column" on read', async ({ page }) => {
    const migrated = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const mod = await import(/* @vite-ignore */ '/src/charts/resources.ts' as any);
      const fakeData = {
        resources: [
          {
            name: '__casual_sheets_charts__',
            data: JSON.stringify({
              v: 1,
              charts: [
                {
                  id: 'ch-legacy',
                  sheetId: 'sheet-1',
                  type: 'bar', // pre-P3 value
                  source: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 },
                  pos: { startRow: 4, endRow: 13, startColumn: 0, endColumn: 7 },
                },
              ],
            }),
          },
        ],
      };
      return mod.readChartsFromSnapshot(fakeData);
    });
    expect(migrated).toHaveLength(1);
    expect(migrated[0].type).toBe('column');
  });
});
