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
import { readStyle, selectRange, waitForUniver } from './_helpers';

test.describe('Undo / Redo (Quick Access Toolbar)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await selectRange(page, 'A1');
  });

  test('Bold then Undo via QAT reverts the style', async ({ page }) => {
    await page.getByTestId('ribbon-btn-bold').click();
    let style = (await readStyle(page, 'A1')) as { bl?: 0 | 1 } | null;
    expect(style?.bl).toBe(1);

    await page.getByTestId('qat-undo').click();
    style = (await readStyle(page, 'A1')) as { bl?: 0 | 1 } | null;
    expect(style?.bl ?? 0).toBe(0);
  });

  test('Redo after undo re-applies', async ({ page }) => {
    await page.getByTestId('ribbon-btn-bold').click();
    await page.getByTestId('qat-undo').click();
    await page.getByTestId('qat-redo').click();
    const style = (await readStyle(page, 'A1')) as { bl?: 0 | 1 } | null;
    expect(style?.bl).toBe(1);
  });
});

test.describe('Strikethrough', () => {
  test('Strike button toggles st on the cell', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await selectRange(page, 'A1');

    const strike = page.getByTestId('ribbon-btn-strikethrough');
    await expect(strike).toHaveAttribute('aria-pressed', 'false');
    await strike.click();
    await expect(strike).toHaveAttribute('aria-pressed', 'true');

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const style = (await readStyle(page, 'A1')) as any;
    expect(style?.st?.s).toBe(1);
  });
});

test.describe('Number format dropdown + decimals', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await selectRange(page, 'A1');
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 1234.5 });
    });
  });

  test('Selecting "Number (2 decimals)" applies #,##0.00', async ({ page }) => {
    await page.getByTestId('ribbon-select-num-format').selectOption('number');
    const style = (await readStyle(page, 'A1')) as { n?: { pattern: string } } | null;
    expect(style?.n?.pattern).toBe('#,##0.00');
  });

  test('Increase decimals (Format menu) adds a decimal place', async ({ page }) => {
    // Start from Number format with 2 decimals; bump to 3.
    await page.getByTestId('ribbon-select-num-format').selectOption('number');
    // Decimal +/- live in the Format menu now.
    await page.getByTestId('menubar-format').click();
    await page.getByTestId('menu-item-decimal-up').click();
    const style = (await readStyle(page, 'A1')) as { n?: { pattern: string } } | null;
    expect(style?.n?.pattern).toMatch(/0\.0{3,}/);
  });
});

test.describe('Format Painter', () => {
  test('Painter button is reachable and clickable', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    // Just verify the button exists and is enabled — full painter interaction
    // requires canvas mouse work that's brittle in headless.
    await expect(page.getByTestId('ribbon-btn-format-painter')).toBeEnabled();
  });
});

test.describe('Find & Replace (Edit menu)', () => {
  test('Edit → Find & Replace dispatches the command', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    const fired = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      return new Promise<boolean>((resolve) => {
        const log: string[] = [];
        api.addEvent(api.Event.CommandExecuted, (e) => {
          const id = (e as { id?: string }).id ?? '';
          log.push(id);
        });
        document.querySelector<HTMLButtonElement>('[data-testid="menubar-edit"]')?.click();
        setTimeout(() => {
          document
            .querySelector<HTMLButtonElement>('[data-testid="menu-item-find-replace"]')
            ?.click();
        }, 50);
        setTimeout(
          () =>
            resolve(
              log.some((id) => id.includes('find-dialog') || id.includes('find-replace')),
            ),
          1500,
        );
      });
    });
    expect(fired).toBe(true);
  });
});

test.describe('Clipboard buttons (no external clipboard write — verify command dispatch)', () => {
  test('Copy ribbon button dispatches univer.command.copy', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await selectRange(page, 'A1');

    const fired = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      return new Promise<boolean>((resolve) => {
        const sub = api.addEvent(api.Event.CommandExecuted, (e) => {
          if ((e as { id?: string }).id === 'univer.command.copy') {
            sub.dispose();
            resolve(true);
          }
        });
        document
          .querySelector<HTMLButtonElement>('[data-testid="ribbon-btn-copy"]')
          ?.click();
        setTimeout(() => resolve(false), 1500);
      });
    });
    expect(fired).toBe(true);
  });
});
