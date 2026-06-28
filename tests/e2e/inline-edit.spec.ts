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
import { mainCanvas, readCell, waitForUniver } from './_helpers';

/**
 * Phase 1.1.1 — inline cell editing works.
 * The cell editor's `FormulaEditor` component is registered by
 * `@univerjs/sheets-formula-ui` (NOT `sheets-formula`). Forgetting it leaves
 * the grid clickable but uneditable — this suite is the canary.
 */

test.describe('Inline cell editing', () => {
  test('Double-click on a cell opens the inline editor and types commit on Enter', async ({
    page,
  }) => {
    await page.goto('/');
    await waitForUniver(page);

    const grid = mainCanvas(page);
    const box = await grid.boundingBox();
    expect(box).not.toBeNull();
    if (!box) return;

    // Aim at roughly B2 — inside the data area, past row/column headers.
    await page.mouse.dblclick(box.x + 200, box.y + 60);

    // The cell editor is a `[data-u-comp="editor"]` contenteditable. Univer
    // animates it from off-screen, so we don't assert visibility — we assert
    // that keys typed after the dblclick land in the targeted cell.
    await page.keyboard.type('inline-works');
    await page.keyboard.press('Enter');

    // Default column widths put roughly (200, 60) inside B2.
    const cell = await readCell(page, 'B2');
    expect(cell?.v).toBe('inline-works');
  });

  // Note: "type-while-selected" is covered by the F2 test in text-edit.spec.ts —
  // it's the same code path on the Univer side (CellEditVisible operation),
  // and F2 doesn't depend on canvas keyboard-focus race conditions in headless
  // Playwright.
});
