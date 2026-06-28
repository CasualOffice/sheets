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
import { waitForUniver } from './_helpers';

/**
 * Excel-paste merge preservation. The hook lives in
 * `apps/web/src/univer/paste-merge-hook.ts` and is registered via
 * the FUniver facade's injector once Univer mounts.
 *
 * We can't reliably dispatch a synthetic `ClipboardEvent` carrying
 * HTML through Univer's listener in headless Chromium (the system
 * clipboard is what Univer reads, not the dispatched event). So the
 * hook also self-exposes its `onPasteCells` callback on
 * `window.__pasteMergeHook__` in dev — these tests drive that
 * directly with synthetic cell-matrix shapes. Production builds
 * tree-shake the exposure via `import.meta.env.DEV`.
 */

async function callHook(
  page: Page,
  data: Array<[number, number, { rowSpan?: number; colSpan?: number }]>,
) {
  return page.evaluate(
    ({ data }) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const hook = (window as any).__pasteMergeHook__ as
        | ((
            from: unknown,
            to: { range: { rows: number[]; cols: number[] }; unitId: string; subUnitId: string },
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            matrix: { forValue: (cb: (r: number, c: number, cell: any) => void) => void },
            payload?: unknown,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
          ) => { undos: any[]; redos: any[] })
        | undefined;
      if (!hook) return { found: false, mutations: null as unknown };

      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      const pasteTo = {
        range: { rows: [10, 11, 12, 13], cols: [5, 6, 7, 8] },
        unitId: wb.getId(),
        subUnitId: wb.getActiveSheet()!.getSheetId(),
      };
      const matrix = {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        forValue(cb: (r: number, c: number, cell: any) => void) {
          for (const [r, c, cell] of data) cb(r, c, cell);
        },
      };
      const result = hook(null, pasteTo, matrix, {});
      return { found: true, mutations: result };
    },
    { data },
  );
}

test.describe('Excel-paste merge hook', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('2x2 colspan/rowspan cell emits add-worksheet-merge with correct range', async ({ page }) => {
    const r = await callHook(page, [[0, 0, { rowSpan: 2, colSpan: 2 }]]);
    expect(r.found).toBe(true);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const m = r.mutations as { redos: any[]; undos: any[] };
    expect(m.redos).toHaveLength(1);
    expect(m.redos[0].id).toBe('sheet.mutation.add-worksheet-merge');
    expect(m.redos[0].params.ranges).toHaveLength(1);
    expect(m.redos[0].params.ranges[0]).toMatchObject({
      startRow: 10,
      endRow: 11,
      startColumn: 5,
      endColumn: 6,
    });
    expect(m.undos[0].id).toBe('sheet.mutation.remove-worksheet-merge');
  });

  test('cells without spans produce no mutations', async ({ page }) => {
    const r = await callHook(page, [
      [0, 0, {}],
      [0, 1, { colSpan: 1 }],
      [1, 0, { rowSpan: 1, colSpan: 1 }],
    ]);
    expect(r.found).toBe(true);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const m = r.mutations as { redos: any[]; undos: any[] };
    expect(m.redos).toEqual([]);
    expect(m.undos).toEqual([]);
  });

  test('col-width hook emits set-worksheet-col-width with parsed widths', async ({ page }) => {
    const r = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const hook = (window as any).__pasteColWidthHook__ as
        | ((
            to: { range: { rows: number[]; cols: number[] }; unitId: string; subUnitId: string },
            colProperties: Record<string, string>[],
            payload: unknown,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
          ) => { undos: any[]; redos: any[] })
        | undefined;
      if (!hook) return { found: false };
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      const pasteTo = {
        range: { rows: [0], cols: [5, 6, 7] },
        unitId: wb.getId(),
        subUnitId: wb.getActiveSheet()!.getSheetId(),
      };
      const colProperties = [
        { width: '120' },
        { width: '60px' }, // strips px
        { width: '' }, // skipped (empty)
      ];
      return { found: true, mutations: hook(pasteTo, colProperties, {}) };
    });
    expect(r.found).toBe(true);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const m = (r as any).mutations as { redos: any[] };
    expect(m.redos).toHaveLength(1);
    expect(m.redos[0].id).toBe('sheet.mutation.set-worksheet-col-width');
    const params = m.redos[0].params;
    // Only the two columns with valid widths should land.
    expect(params.colWidth).toEqual({ 5: 120, 6: 60 });
    expect(params.ranges[0].startColumn).toBe(5);
    expect(params.ranges[0].endColumn).toBe(6);
  });

  test('row-height hook emits set-worksheet-row-height with parsed heights', async ({ page }) => {
    const r = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const hook = (window as any).__pasteRowHeightHook__ as
        | ((
            to: { range: { rows: number[]; cols: number[] }; unitId: string; subUnitId: string },
            rowProperties: Record<string, string>[],
            payload: unknown,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
          ) => { undos: any[]; redos: any[] })
        | undefined;
      if (!hook) return { found: false };
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      const pasteTo = {
        range: { rows: [3, 4, 5], cols: [0] },
        unitId: wb.getId(),
        subUnitId: wb.getActiveSheet()!.getSheetId(),
      };
      const rowProperties = [
        { height: '40' },
        { height: '24px' }, // strips px
        { height: 'auto' }, // skipped
      ];
      return { found: true, mutations: hook(pasteTo, rowProperties, {}) };
    });
    expect(r.found).toBe(true);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const m = (r as any).mutations as { redos: any[] };
    expect(m.redos).toHaveLength(1);
    expect(m.redos[0].id).toBe('sheet.mutation.set-worksheet-row-height');
    expect(m.redos[0].params.rowHeight).toEqual({ 3: 40, 4: 24 });
    expect(m.redos[0].params.ranges[0].startRow).toBe(3);
    expect(m.redos[0].params.ranges[0].endRow).toBe(4);
  });

  test('col-width hook is a no-op when nothing parseable comes in', async ({ page }) => {
    const r = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const hook = (window as any).__pasteColWidthHook__ as
        | ((
            to: { range: { rows: number[]; cols: number[] }; unitId: string; subUnitId: string },
            colProperties: Record<string, string>[],
            payload: unknown,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
          ) => { undos: any[]; redos: any[] })
        | undefined;
      if (!hook) return { found: false };
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      const pasteTo = {
        range: { rows: [0], cols: [5, 6] },
        unitId: wb.getId(),
        subUnitId: wb.getActiveSheet()!.getSheetId(),
      };
      return {
        found: true,
        mutations: hook(pasteTo, [{ width: 'auto' }, { width: '' }], {}),
      };
    });
    expect(r.found).toBe(true);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((r as any).mutations.redos).toEqual([]);
  });

  test('multiple merges in one paste batch into a single mutation', async ({ page }) => {
    const r = await callHook(page, [
      [0, 0, { colSpan: 2 }],
      [1, 0, { rowSpan: 2 }],
      [2, 1, { colSpan: 3 }],
    ]);
    expect(r.found).toBe(true);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const m = r.mutations as { redos: any[] };
    expect(m.redos).toHaveLength(1);
    expect(m.redos[0].params.ranges).toHaveLength(3);
    // Verify each range's geometry — paste origin is rows[10..13], cols[5..8].
    // First merge: row 0, col 0, colSpan 2 → 1x2 at (10, 5..6).
    expect(m.redos[0].params.ranges[0]).toMatchObject({
      startRow: 10,
      endRow: 10,
      startColumn: 5,
      endColumn: 6,
    });
    // Second: row 1, col 0, rowSpan 2 → 2x1 at (11..12, 5).
    expect(m.redos[0].params.ranges[1]).toMatchObject({
      startRow: 11,
      endRow: 12,
      startColumn: 5,
      endColumn: 5,
    });
    // Third: row 2, col 1, colSpan 3 → 1x3 at (12, 6..8).
    expect(m.redos[0].params.ranges[2]).toMatchObject({
      startRow: 12,
      endRow: 12,
      startColumn: 6,
      endColumn: 8,
    });
  });
});
