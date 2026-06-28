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
 * Dynamic array formulas + spill — Phase 4, T4.1.
 *
 * Univer's formula engine evaluates modern dynamic-array functions (SEQUENCE,
 * UNIQUE, TRANSPOSE, FILTER, SORT, XLOOKUP…) and spills the result into the
 * neighbouring cells automatically — the anchor holds the formula, the spilled
 * cells hold values, and a blocked spill yields `#SPILL!`. This capability
 * shipped with the 0.25 fork but had no coverage; this locks it against
 * regressions (e.g. a future fork bump silently dropping spill).
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyWs = any;

async function read(page: import('@playwright/test').Page, a1: string) {
  return page.evaluate((ref) => {
    const api = window.__univerAPI!;
    const ws = api.getActiveWorkbook()!.getActiveSheet() as AnyWs;
    const c = ws.getRange(ref);
    return { v: c.getValue?.() ?? null, f: c.getFormula?.() ?? '' };
  }, a1);
}

test('SEQUENCE spills down; the anchor holds the formula, spilled cells hold values', async ({
  page,
}) => {
  test.setTimeout(60_000);
  await page.goto('/');
  await waitForUniver(page);

  await page.evaluate(() => {
    const ws = window.__univerAPI!.getActiveWorkbook()!.getActiveSheet() as AnyWs;
    ws.getRange('A1').setValue({ f: '=SEQUENCE(3,1)' });
  });

  await expect.poll(async () => (await read(page, 'A3')).v).toBe(3);
  expect((await read(page, 'A1')).v).toBe(1);
  expect((await read(page, 'A2')).v).toBe(2);
  // Only the anchor carries the formula; spilled cells are plain values.
  expect((await read(page, 'A1')).f).toContain('SEQUENCE');
  expect((await read(page, 'A2')).f).toBe('');
});

test('UNIQUE dedups + spills, and TRANSPOSE spills across columns', async ({ page }) => {
  test.setTimeout(60_000);
  await page.goto('/');
  await waitForUniver(page);

  await page.evaluate(() => {
    const ws = window.__univerAPI!.getActiveWorkbook()!.getActiveSheet() as AnyWs;
    ws.getRange('A1').setValue({ v: 'x' });
    ws.getRange('A2').setValue({ v: 'x' });
    ws.getRange('A3').setValue({ v: 'y' });
    ws.getRange('C1').setValue({ f: '=UNIQUE(A1:A3)' });
    ws.getRange('G1').setValue({ f: '=TRANSPOSE(A1:A3)' });
  });

  await expect.poll(async () => (await read(page, 'C2')).v).toBe('y');
  expect((await read(page, 'C1')).v).toBe('x'); // x, y (deduped)
  // TRANSPOSE: column A1:A3 → row G1:I1.
  await expect.poll(async () => (await read(page, 'I1')).v).toBe('y');
  expect((await read(page, 'G1')).v).toBe('x');
  expect((await read(page, 'H1')).v).toBe('x');
});

test('a blocked spill yields #SPILL! and leaves the blocking cell intact', async ({ page }) => {
  test.setTimeout(60_000);
  await page.goto('/');
  await waitForUniver(page);

  await page.evaluate(() => {
    const ws = window.__univerAPI!.getActiveWorkbook()!.getActiveSheet() as AnyWs;
    ws.getRange('E2').setValue({ v: 'blocker' });
    ws.getRange('E1').setValue({ f: '=SEQUENCE(3,1)' });
  });

  await expect.poll(async () => (await read(page, 'E1')).v).toBe('#SPILL!');
  expect((await read(page, 'E2')).v).toBe('blocker'); // not overwritten
});
