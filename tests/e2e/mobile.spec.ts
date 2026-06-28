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

// Phone-sized viewport. Drives the @media (max-width: 480px) / 720px
// rules in styles.css.
const PHONE_VIEWPORT = { width: 375, height: 667 };

/**
 * Mobile shell — viewer + light-editor lane per CLAUDE.md's mobile
 * scope. Univer's canvas owns its own touch gestures; the React shell's
 * job is making sure the chrome around the canvas reads and behaves
 * properly on a phone.
 *
 * NOT a "full mobile editing" suite — chart insert, pivot field-list,
 * complex formula composition are explicitly out-of-scope on mobile.
 * These tests lock the surfaces a viewer + casual editor needs.
 */
test.describe('Mobile shell (375 × 667)', () => {
  test.beforeEach(async ({ page }) => {
    await page.setViewportSize(PHONE_VIEWPORT);
  });

  // Dismiss the home overlay (added in the home-gallery commit) so the
  // editor chrome behind it is reachable. The home is the right default
  // for real users but it covers menu bar / toolbar / formula bar for
  // the chrome assertions below.
  async function dismissHome(page: import('@playwright/test').Page) {
    await page.goto('/');
    await waitForUniver(page);
    const home = page.getByTestId('home-screen');
    if (await home.count()) {
      await page.getByTestId('home-close').click();
      await expect(home).toHaveCount(0);
    }
  }

  test('renders without horizontal overflow', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    // Document width may slightly exceed viewport because the menubar +
    // toolbar are designed to horizontally scroll on phones. The BODY
    // itself shouldn't overflow.
    const bodyOverflow = await page.evaluate(() => {
      const body = document.body;
      return body.scrollWidth - body.clientWidth;
    });
    expect(bodyOverflow, `body overflowed by ${bodyOverflow}px`).toBeLessThanOrEqual(2);
  });

  test('menu bar is reachable and scrolls horizontally to reveal all items', async ({ page }) => {
    await dismissHome(page);

    // Every menubar item must still exist in the DOM (we don't drop any
    // on mobile — they're scrollable, not hidden).
    for (const id of ['file', 'edit', 'view', 'insert', 'format', 'data', 'help']) {
      await expect(page.getByTestId(`menubar-${id}`)).toBeAttached();
    }

    // File menu still opens and shows its core items.
    await page.getByTestId('menubar-file').click();
    await expect(page.getByTestId('menu-item-open')).toBeVisible();
    await expect(page.getByTestId('menu-item-save')).toBeVisible();
  });

  test('toolbar stays visible (compact) on a 375 px viewport', async ({ page }) => {
    await dismissHome(page);
    // Mobile-pass change (was hidden previously): the toolbar is now
    // a single-row horizontal-scroll strip so light formatting is one
    // tap away. Row 2 of every ribbon group is hidden via CSS.
    await expect(page.getByTestId('toolbar')).toBeVisible();
    await expect(page.getByTestId('formula-bar')).toBeVisible();
    await expect(page.getByTestId('grid-host')).toBeVisible();
  });

  test('formula-bar input is ≥ 16 px so iOS Safari does not focus-zoom', async ({ page }) => {
    await dismissHome(page);
    const input = page.locator('.formula-bar__input').first();
    await expect(input).toBeAttached();
    const fontSize = await input.evaluate(
      (el) => parseFloat(window.getComputedStyle(el).fontSize),
    );
    // iOS Safari triggers an auto-zoom on inputs with computed font-size
    // < 16 px. The mobile breakpoint pins this; if it ever regresses,
    // the whole layout shifts on focus and editing becomes unusable.
    expect(fontSize, `formula-bar input font-size was ${fontSize}px`).toBeGreaterThanOrEqual(16);
  });

  test('bottom action bar is visible on mobile with formatting buttons', async ({ page }) => {
    await dismissHome(page);
    // The bar is CSS-hidden above 480 px so it MUST be visible at our
    // 375 × 667 viewport. Tap targets cover B/I/U + $/% + decimal +
    // align — the most-used formatters for a thumb-driven mobile edit.
    const bar = page.getByTestId('mobile-action-bar');
    await expect(bar).toBeVisible();

    for (const id of [
      'mobile-bar-bold',
      'mobile-bar-italic',
      'mobile-bar-underline',
      'mobile-bar-currency',
      'mobile-bar-percent',
      'mobile-bar-align-left',
      'mobile-bar-align-center',
      'mobile-bar-align-right',
    ]) {
      await expect(page.getByTestId(id), `${id} attached`).toBeAttached();
    }
  });

  test('touch-drag on the canvas dispatches wheel events (TouchPanDriver wired)', async ({ page }) => {
    // Univer 0.24 has no native touch-pan — we synthesize wheel events
    // from pointermove deltas in `useTouchPan`. Dismiss the home,
    // instrument the canvas to count wheel events, simulate a single-
    // finger vertical swipe via pointer events, and assert the wheel
    // counter went up. Uses pointer events (not touch) because Univer's
    // own drag-to-select listens on pointermove — the hook has to
    // block the SAME stream Univer reads, otherwise the swipe extends
    // a phantom cell selection.
    await dismissHome(page);

    // Wait for the visible render-canvas. Univer mounts two canvases —
    // a visible main one + a 0×0 print one — so filter to the main.
    await page.waitForFunction(
      () => {
        const all = Array.from(document.querySelectorAll('canvas[data-u-comp="render-canvas"]'));
        return all.some((c) => (c as HTMLCanvasElement).width > 0);
      },
      { timeout: 10_000 },
    );

    // Instrument the canvas: count wheel events.
    await page.evaluate(() => {
      const all = Array.from(document.querySelectorAll('canvas[data-u-comp="render-canvas"]'));
      const c = all.find((el) => (el as HTMLCanvasElement).width > 0) as HTMLCanvasElement | undefined;
      if (!c) return;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (window as any).__wheelCount = 0;
      c.addEventListener(
        'wheel',
        () => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__wheelCount = ((window as any).__wheelCount ?? 0) + 1;
        },
        true,
      );
    });

    const box = await page.locator('[data-testid="univer-host"]').boundingBox();
    expect(box, 'univer-host bbox').not.toBeNull();
    if (!box) return;

    const cx = box.x + box.width / 2;
    const startY = box.y + box.height - 60;
    const endY = box.y + 80;

    // Drive raw PointerEvents (with pointerType: 'touch') in capture
    // order so useTouchPan's document-level listener gets them.
    for (let step = 0; step <= 4; step++) {
      const y = startY - ((startY - endY) * (step / 4));
      await page.evaluate(
        ({ x, y, step }) => {
          const targetEl = document.elementFromPoint(x, y) ?? document.body;
          const type = step === 0 ? 'pointerdown' : step === 4 ? 'pointerup' : 'pointermove';
          const ev = new PointerEvent(type, {
            pointerId: 1,
            pointerType: 'touch',
            bubbles: true,
            cancelable: true,
            clientX: x,
            clientY: y,
            isPrimary: true,
            button: step === 4 ? -1 : 0,
            buttons: step === 4 ? 0 : 1,
          });
          targetEl.dispatchEvent(ev);
        },
        { x: cx, y, step },
      );
      await page.waitForTimeout(40);
    }

    // useTouchPan dispatches one wheel per pointermove past the
    // threshold — we ran 3 moves, so we expect at least 2 wheels.
    const wheelCount = await page.evaluate(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__wheelCount as number | undefined,
    );
    expect(wheelCount ?? 0, 'touch-drag should have dispatched wheel events on the canvas')
      .toBeGreaterThanOrEqual(2);
  });
});

test.describe('Collab indicator', () => {
  test('renders in solo mode by default', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    const indicator = page.getByTestId('collab-indicator');
    await expect(indicator).toBeVisible();
    await expect(indicator).toHaveAttribute('data-collab-status', 'off');
    await expect(indicator).toContainText(/Solo/i);
  });
});
