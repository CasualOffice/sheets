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
 * Offline fonts (desktop slice 4a). The Tauri app has no network, so the
 * sheets bootstrap declares Inter + Material Symbols from locally-bundled woff2
 * via injected @font-face — but ONLY in desktop mode (`?desk=1`). The web build
 * keeps the Google Fonts CDN (per the design-system rule) and gains zero bytes.
 */

test('desktop injects local @font-face for the chrome fonts', async ({ page }) => {
  await page.goto('/?desk=1');
  await waitForUniver(page);
  const css = await page.evaluate(
    () => document.getElementById('__deskapp_fonts__')?.textContent ?? '',
  );
  expect(css).toContain('Material Symbols Outlined');
  expect(css).toContain('inter-400.woff2');
  expect(css).toContain('./fonts/material-symbols-outlined.subset.woff2');
});

test('web build does NOT inject local fonts and keeps the Google Fonts CDN', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  // No local @font-face block in web.
  await expect(page.locator('#__deskapp_fonts__')).toHaveCount(0);
  // The CDN stylesheet link is still present.
  const cdn = await page.locator('link[href*="fonts.googleapis.com"]').count();
  expect(cdn).toBeGreaterThan(0);
});
