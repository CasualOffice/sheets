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

/**
 * Paste-to-fit grid growth — Phase 1, T1.3 (docs/COMPETITIVE_ROADMAP.md).
 *
 * A paste whose destination exceeds the current grid extent must extend the
 * sheet so the overflow is visible/valid (Univer's clipboard service expands the
 * destination un-clamped, so the rows land beyond `rowCount`). `growToFitPaste`
 * (registered as the `onBeforePaste` clipboard hook) grows the sheet to cover the
 * destination before the paste mutations apply.
 *
 * We drive the hook callback directly via `__pasteGrowthHook__`: a real Ctrl+V /
 * clipboard paste doesn't reach Univer's listener in headless Chromium (the same
 * constraint `paste-merge-hook.ts` documents and tests around). The real paste
 * path passes this same expanded range to the hook — see `_getPastedRange` /
 * `_expandOrShrinkRowsCols` in the fork's clipboard.service.ts (un-clamped).
 */
test('paste destination beyond the grid extent grows the sheet to fit', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  const ROWS = 2000;

  const before = await page.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (window.__univerAPI!.getActiveWorkbook()!.getActiveSheet() as any).getMaxRows() as number;
  });
  expect(before).toBeLessThan(ROWS); // grid starts smaller than the paste

  const result = await page.evaluate((n) => {
    const api = window.__univerAPI!;
    const wb = api.getActiveWorkbook()!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws = wb.getActiveSheet() as any;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const hook = (globalThis as any).__pasteGrowthHook__ as
      | ((p: { unitId: string; subUnitId: string; range: { rows: number[]; cols: number[] } }) => boolean)
      | undefined;
    if (!hook) return { error: 'hook-not-exposed' as const };

    const rows = Array.from({ length: n }, (_, i) => i); // [0..n-1] ascending, like a real paste
    const ret = hook({ unitId: wb.getId(), subUnitId: ws.getSheetId(), range: { rows, cols: [0] } });

    // After growth, a write at the deep row (out-of-bounds before) must succeed.
    let deep: unknown = null;
    let writeError: string | null = null;
    try {
      ws.getRange(n - 1, 0).setValue('deep');
      deep = ws.getRange(n - 1, 0).getValue();
    } catch (e) {
      writeError = (e as Error).message;
    }
    return { ret, after: ws.getMaxRows() as number, deep, writeError };
  }, ROWS);

  expect(result.error).toBeUndefined();
  expect(result.ret).toBe(true); // never blocks the paste
  expect(result.after).toBeGreaterThanOrEqual(ROWS); // sheet grew to fit
  expect(result.writeError).toBeNull(); // deep cell is now in-bounds
  expect(result.deep).toBe('deep'); // and holds its value
});
