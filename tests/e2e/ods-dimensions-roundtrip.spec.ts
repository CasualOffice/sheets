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
import { waitForUniver } from './_helpers';

type OdsModule = typeof import('../../apps/web/src/ods');
type OdsWorkbookData = Parameters<OdsModule['workbookDataToOds']>[0];

declare global {
  interface Window {
    __odsDims?: typeof import('../../apps/web/src/ods');
  }
}

test.describe('ods dimensions round-trip', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const mod = await import(/* @vite-ignore */ '/src/ods/index.ts' as any);
      window.__odsDims = mod;
    });
  });

  test('snapshot row heights and column widths survive ods export and import', async ({ page }) => {
    const result = await page.evaluate(async () => {
      const snapshot: OdsWorkbookData = {
        id: 'wb-dims-1',
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
            columnData: {
              0: { w: 125 },
              1: { w: 77 },
            },
            rowData: {
              0: { h: 24 },
              1: { h: 18 },
            },
          },
        },
      };

      const blob = await window.__odsDims!.workbookDataToOds(snapshot);
      const buf = await blob.arrayBuffer();
      const reloaded = await window.__odsDims!.odsToWorkbookData(buf);
      const sheetId = reloaded.sheetOrder[0];
      const sheet = reloaded.sheets[sheetId];
      return {
        col0: sheet?.columnData?.[0]?.w ?? null,
        col1: sheet?.columnData?.[1]?.w ?? null,
        row0: sheet?.rowData?.[0]?.h ?? null,
        row1: sheet?.rowData?.[1]?.h ?? null,
      };
    });

    expect(result.col0).toBe(125);
      expect(result.col1).toBe(77);
      expect(result.row0).toBe(24);
      expect(result.row1).toBe(18);
  });

  test('hidden row and column metadata survive ods export and import', async ({ page }) => {
    const result = await page.evaluate(async () => {
      const snapshot: OdsWorkbookData = {
        id: 'wb-dims-hidden-1',
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
            columnData: {
              0: { w: 125 },
              1: { hd: 1 },
            },
            rowData: {
              0: { h: 24 },
              1: { hd: 1 },
            },
          },
        },
      };

      const blob = await window.__odsDims!.workbookDataToOds(snapshot);
      const buf = await blob.arrayBuffer();
      const XLSX = await import('/node_modules/.vite/deps/@e965_xlsx.js');
      const cfb = XLSX.CFB.read(new Uint8Array(buf), { type: 'array' });
      const idx = cfb.FullPaths.indexOf('Root Entry/content.xml');
      const contentXml = new TextDecoder().decode(cfb.FileIndex[idx]?.content as Uint8Array);
      const reloaded = await window.__odsDims!.odsToWorkbookData(buf);
      const sheetId = reloaded.sheetOrder[0];
      const sheet = reloaded.sheets[sheetId];
      return {
        col1Hidden: sheet?.columnData?.[1]?.hd ?? null,
        row1Hidden: sheet?.rowData?.[1]?.hd ?? null,
        contentXml,
      };
    });

    expect(result.col1Hidden).toBe(1);
    expect(result.row1Hidden).toBe(1);
    expect(result.contentXml).toContain('table:visibility="collapse"');
  });
});
