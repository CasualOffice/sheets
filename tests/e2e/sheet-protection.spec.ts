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
 * Per-sheet protection — Phase 4, T4.4 (slice 2).
 *
 * "Data → Protect sheet" marks the active worksheet protected via Univer's
 * worksheet-permission facade (collab model: other editors can't change it; the
 * protector still can — the chosen Google-style model). It's worksheet-scoped,
 * so sibling sheets are unaffected. Toggling the menu item lifts it.
 *
 * Asserts `isProtected()` — the faithful signal for this model (`canEditCell`
 * stays true for the local protector by design).
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyApi = any;

const isProtected = (page: import('@playwright/test').Page) =>
  page.evaluate(() => {
    const api = window.__univerAPI as AnyApi;
    return api.getActiveWorkbook().getActiveSheet().getWorksheetPermission().isProtected() === true;
  });

test('Data → Protect sheet toggles worksheet protection on the active sheet', async ({ page }) => {
  test.setTimeout(60_000);
  await page.goto('/');
  await waitForUniver(page);

  expect(await isProtected(page)).toBe(false);

  // Protect via the menu.
  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-protect-worksheet').click();
  await expect.poll(() => isProtected(page)).toBe(true);

  // Toggle off via the menu (label now shows the protected state).
  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-protect-worksheet').click();
  await expect.poll(() => isProtected(page)).toBe(false);
});
