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

/**
 * Command palette — UX_AUDIT.md §4.2 / Phase 4 #15.
 *
 * Sheet already had Alt+Q ("Search / Tell Me", Office convention).
 * This spec asserts the new Ctrl+Shift+P alias (VS Code / Linear /
 * Notion convention) opens the same CommandSearchDialog so users
 * coming from those tools find it on first try.
 *
 * Ctrl+K is intentionally left bound to Insert Link (Excel
 * convention) — both chords coexist, neither is repurposed.
 */
import { test, expect } from '@playwright/test';
import { waitForUniver } from './_helpers';

test.describe('Command palette chord', () => {
  test('Ctrl+Shift+P opens CommandSearchDialog', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.keyboard.press('Control+Shift+P');
    // CommandSearchDialog has a searchable input and a results list —
    // the dialog testid is the canonical handle.
    await expect(page.getByTestId('command-search-dialog')).toBeVisible();
  });

  test('Alt+Q still opens the same dialog (legacy / Office chord)', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.keyboard.press('Alt+Q');
    await expect(page.getByTestId('command-search-dialog')).toBeVisible();
  });
});
