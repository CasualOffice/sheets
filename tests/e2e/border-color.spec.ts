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
import { selectRange, waitForUniver } from './_helpers';

/**
 * Border color picker. The Borders split-button in the Home ribbon exposes a
 * "Line color" row at the bottom of its dropdown. Picking a swatch updates
 * the per-session color; the next border style click applies with that color.
 *
 * Lock-down test: previously the border color was hardcoded #666666 with no
 * way to change it. If the dropdown loses the color row or BordersControl
 * stops threading the color through to `setBorders`, casual users can't pick
 * a border color anymore.
 */
test('Borders dropdown applies the picked color to the active range', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  await selectRange(page, 'B2');

  // Open the borders dropdown, pick a non-default swatch (Excel red).
  await page.getByTestId('ribbon-dropdown-borders-caret').click();
  await page.getByTestId('ribbon-dropdown-borders-color-row').waitFor();

  // Quick DOM probe so we fail fast if the items map regresses.
  const itemIds = await page.evaluate(() =>
    Array.from(document.querySelectorAll('[data-testid^="ribbon-dropdown-borders-item-"]')).map(
      (el) => (el as HTMLElement).dataset.testid,
    ),
  );
  expect(itemIds).toEqual(
    expect.arrayContaining([
      'ribbon-dropdown-borders-item-all',
      'ribbon-dropdown-borders-item-outside',
      'ribbon-dropdown-borders-item-top',
      'ribbon-dropdown-borders-item-bottom',
      'ribbon-dropdown-borders-item-left',
      'ribbon-dropdown-borders-item-right',
      'ribbon-dropdown-borders-item-none',
    ]),
  );

  await page.getByTestId('ribbon-dropdown-borders-color-d93025').click();

  // Now apply the All-borders style.
  await page.getByTestId('ribbon-dropdown-borders-item-all').click();

  // The cell's style must carry the red color on every border side.
  const border = await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook()!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = wb.getActiveSheet();
    const cd = ws.getRange('B2').getCellData();
    const styleRef = cd?.s;
    const style =
      typeof styleRef === 'string'
        ? wb.getWorkbook().getStyles().get(styleRef)
        : (styleRef ?? null);
    return style?.bd ?? null;
  });

  expect(border, 'expected border metadata on the cell').toBeTruthy();
  const expected = '#d93025';
  for (const side of ['t', 'b', 'l', 'r']) {
    expect(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (border as any)[side]?.cl?.rgb?.toLowerCase(),
      `${side} border color`,
    ).toBe(expected);
  }
});
