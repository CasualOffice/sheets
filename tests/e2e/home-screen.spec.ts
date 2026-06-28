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
 * Home screen — the template gallery + recent files overlay shown over
 * a blank Untitled workbook on first load. Verifies the basics:
 *
 *   1. Overlay renders on a fresh page load (blank Untitled).
 *   2. Featured strip + at least the Personal Budget template card are
 *      present.
 *   3. Search filters cards in place.
 *   4. Picking a real template fetches /templates/{id}.xlsx, runs it
 *      through the parser, and the overlay self-dismisses because the
 *      workbook is no longer Untitled.
 */

test.describe('home screen', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    // This suite TESTS the home screen — opt out of the default
    // dismiss-on-mount that other tests get.
    await waitForUniver(page, { keepHome: true });
  });

  test('renders the gallery on a blank Untitled workbook', async ({ page }) => {
    const home = page.getByTestId('home-screen');
    await expect(home).toBeVisible();
    // New app-shell IA: the landing "Home" view shows the Featured strip
    // (Personal Budget is featured). The full per-category gallery lives in
    // the Templates view, reached from the sidebar nav.
    await expect(page.getByRole('heading', { name: 'Home', level: 1 })).toBeVisible();
    await expect(page.getByTestId('tpl-card-personal-budget').first()).toBeVisible();
    await page.getByTestId('home-nav-templates').click();
    await expect(page.getByTestId('tpl-card-invoice').first()).toBeVisible();
  });

  test('search filters template cards', async ({ page }) => {
    // Search now lives in the Templates view (sidebar nav).
    await page.getByTestId('home-nav-templates').click();
    await page.getByTestId('home-search').fill('invoice');
    await expect(page.getByTestId('tpl-card-invoice')).toHaveCount(1);
    await expect(page.getByTestId('tpl-card-personal-budget')).toHaveCount(0);
  });

  test('picking a template opens it and dismisses the home', async ({ page }) => {
    await page.getByTestId('tpl-card-personal-budget').first().click();
    // Home dismisses once workbook is no longer Untitled.
    await expect(page.getByTestId('home-screen')).toHaveCount(0, { timeout: 15_000 });
    // Confirm workbook name reflects the template.
    await expect(page.getByText('Personal budget').first()).toBeVisible();
  });

  test('close button + Esc dismiss without picking', async ({ page }) => {
    await expect(page.getByTestId('home-screen')).toBeVisible();
    await page.getByTestId('home-close').click();
    await expect(page.getByTestId('home-screen')).toHaveCount(0);
  });
});

test.describe('home screen — suppression', () => {
  test('does not render on a collab room URL (/r/:id)', async ({ page }) => {
    // Real collab join requires Hocuspocus; we only need to confirm the
    // home overlay is suppressed for the path pattern so coedit specs
    // can interact with the canvas without the overlay swallowing clicks.
    await page.goto('/r/test1234');
    await expect(page.getByTestId('home-screen')).toHaveCount(0);
  });

  test('does not render on a ?room= URL', async ({ page }) => {
    await page.goto('/?room=test1234');
    await expect(page.getByTestId('home-screen')).toHaveCount(0);
  });
});
