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
    __odsStyles?: typeof import('../../apps/web/src/ods');
  }
}

test.describe('ods style subset round-trip', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const mod = await import(/* @vite-ignore */ '/src/ods/index.ts' as any);
      window.__odsStyles = mod;
    });
  });

  test('font, fill, alignment, and number format survive ods export and import', async ({ page }) => {
    const result = await page.evaluate(async () => {
      const snapshot: OdsWorkbookData = {
        id: 'wb-style-1',
        rev: 1,
        name: 'wb',
        appVersion: '0.22.1',
        locale: 1,
        styles: {
          s0: {
            bl: 1,
            cl: { rgb: '#ff0000' },
            bg: { rgb: '#00ff00' },
            ht: 2,
            vt: 2,
            n: { pattern: '0.00%' },
          },
        },
        sheetOrder: ['s1'],
        sheets: {
          s1: {
            id: 's1',
            name: 'Data',
            cellData: {
              0: {
                0: { v: 0.25, s: 's0' },
              },
            },
            rowCount: 100,
            columnCount: 26,
          },
        },
      };

      const blob = await window.__odsStyles!.workbookDataToOds(snapshot);
      const buf = await blob.arrayBuffer();
      const reloaded = await window.__odsStyles!.odsToWorkbookData(buf);
      const sheetId = reloaded.sheetOrder[0];
      const cell = reloaded.sheets[sheetId]?.cellData?.[0]?.[0];
      const style = cell?.s ? reloaded.styles?.[cell.s] : null;
      return style ?? null;
    });

    expect(result?.bl).toBe(1);
    expect(result?.cl?.rgb).toBe('#ff0000');
    expect(result?.bg?.rgb).toBe('#00ff00');
    expect(result?.ht).toBe(2);
    expect(result?.vt).toBe(2);
    expect(result?.n?.pattern).toBe('0.00%');
  });
});
