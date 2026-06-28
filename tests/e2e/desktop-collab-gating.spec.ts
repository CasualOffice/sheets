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
 * Desktop feature-gating — the desktop (Tauri) build is single-user + local-file,
 * so collab affordances are hidden while everything else stays. Driven by
 * `isDesktop()` (the `?desk=1` flag the Tauri shell appends). These specs prove:
 *   - desktop (`?desk=1`): no Share button, no collab cluster, no "Share for
 *     co-editing" menu item — but History (non-collab) still works;
 *   - web (default): all collab affordances present.
 */

test('desktop build hides collab UI but keeps non-collab features', async ({ page }) => {
  test.setTimeout(60_000);
  await page.goto('/?desk=1');
  await waitForUniver(page);

  // Collab cluster + Share button gone.
  await expect(page.getByTestId('titlebar-collab')).toHaveCount(0);
  await expect(page.getByTestId('titlebar-share')).toHaveCount(0);

  // "Share for co-editing…" absent from the File menu.
  await page.getByTestId('menubar-file').click();
  await expect(page.getByTestId('menu-item-start-room')).toHaveCount(0);
  await page.keyboard.press('Escape');

  // Non-collab features remain — e.g. the History panel rail button.
  await expect(page.getByTestId('panel-rail-history')).toBeVisible();
});

test('web build keeps the collab UI (control)', async ({ page }) => {
  test.setTimeout(60_000);
  await page.goto('/');
  await waitForUniver(page);

  await expect(page.getByTestId('titlebar-collab')).toBeVisible();
  await expect(page.getByTestId('titlebar-share')).toBeVisible();
  await page.getByTestId('menubar-file').click();
  await expect(page.getByTestId('menu-item-start-room')).toBeVisible();
});
