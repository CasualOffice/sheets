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
 * Polish #4 — responsive layout. Locks in the grid-host always
 * fills the available vertical space + the status bar always sits
 * at the bottom (with sheet-tabs directly above it) + the toolbar
 * collapses cleanly on the narrow phone breakpoint.
 *
 * The root bug this guards against: CSS-Grid auto-placement was
 * shifting the grid-host into the wrong track when the toolbar was
 * `display:none` (phone @ 480 px), collapsing the grid to a 26 px
 * sliver. Fix in styles.css uses `grid-template-areas` to pin each
 * chrome row by name regardless of which siblings are present.
 */

const VIEWPORTS = [
  { w: 1920, h: 1080, name: 'desktop-large' },
  { w: 1440, h: 900, name: 'desktop' },
  { w: 1280, h: 800, name: 'laptop' },
  { w: 1024, h: 768, name: 'small-laptop' },
  { w: 900, h: 800, name: 'narrow-laptop' },
  { w: 768, h: 1024, name: 'tablet' },
  { w: 480, h: 800, name: 'phone' },
];

test.describe('Responsive layout — chrome and grid sizing', () => {
  for (const v of VIEWPORTS) {
    test(`grid-host fills the viewport and the status bar sits at the bottom @ ${v.name} (${v.w}x${v.h})`, async ({
      page,
    }) => {
      await page.setViewportSize({ width: v.w, height: v.h });
      await page.goto('/');
      await waitForUniver(page);

      const layout = await page.evaluate(() => {
        const q = (sel: string) => {
          const el = document.querySelector(sel) as HTMLElement | null;
          if (!el) return null;
          const r = el.getBoundingClientRect();
          return {
            w: Math.round(r.width),
            h: Math.round(r.height),
            top: Math.round(r.top),
            bottom: Math.round(r.bottom),
          };
        };
        return {
          app: q('[data-testid="app-shell"]'),
          gridHost: q('.grid-host'),
          sheetTabs: q('.sheet-tabs'),
          statusBar: q('.statusbar'),
          viewportH: window.innerHeight,
        };
      });

      expect(layout.app).not.toBeNull();
      expect(layout.gridHost).not.toBeNull();
      expect(layout.sheetTabs).not.toBeNull();
      expect(layout.statusBar).not.toBeNull();

      // 1. App shell fills the full viewport (100dvh).
      expect(layout.app!.h).toBeGreaterThanOrEqual(layout.viewportH - 1);

      // 2. Grid host gets the bulk of the vertical space — at least
      //    60 % of the viewport at any size. Earlier bug had the host
      //    at 26 px (3% of an 800 px viewport).
      expect(layout.gridHost!.h).toBeGreaterThan(layout.viewportH * 0.6);

      // 3. The status bar is the bottom-most strip (the Phase 4 redesign
      //    split selection stats / zoom / undo out of the sheet-tabs row
      //    into their own strip below it). Its bottom edge sits within
      //    2 px of the viewport bottom.
      expect(layout.statusBar!.bottom).toBeGreaterThanOrEqual(layout.viewportH - 2);
      expect(layout.statusBar!.bottom).toBeLessThanOrEqual(layout.viewportH + 1);

      // 4. Sheet tabs sit directly above the status bar (no gap).
      expect(layout.sheetTabs!.bottom).toBeLessThanOrEqual(layout.statusBar!.top + 1);
      expect(layout.sheetTabs!.bottom).toBeGreaterThanOrEqual(layout.statusBar!.top - 2);
    });
  }

  test('phone breakpoint (<= 480) keeps toolbar visible but compact', async ({ page }) => {
    // Earlier behaviour (Polish #4) hid the toolbar on phone. The
    // mobile-pass rework (see commit 255ba43) reversed that: the
    // toolbar stays visible as a single-row horizontal-scroll strip so
    // light formatting (B/I/U + currency/percent/comma + font picker)
    // stays one tap away. mobile.spec.ts:66 enforces the same. This
    // test pins the height at <=44 px so the chrome stays thumb-
    // reachable above the formula bar.
    await page.setViewportSize({ width: 480, height: 800 });
    await page.goto('/');
    await waitForUniver(page);
    const toolbar = await page.evaluate(() => {
      const el = document.querySelector('.toolbar') as HTMLElement | null;
      const cs = el ? window.getComputedStyle(el) : null;
      return { display: cs?.display, rect: el?.getBoundingClientRect() };
    });
    expect(toolbar.display).not.toBe('none');
    expect(toolbar.rect?.height ?? 0).toBeGreaterThan(0);
    expect(toolbar.rect?.height ?? 0).toBeLessThanOrEqual(48);
  });

  test('grid track for toolbar is compact at phone, full --toolbar-h at wider viewports', async ({
    page,
  }) => {
    // Phone: toolbar collapses to its --toolbar-h (40 px at ≤480).
    await page.setViewportSize({ width: 480, height: 800 });
    await page.goto('/');
    await waitForUniver(page);
    const rowsAtPhone = await page.evaluate(() => {
      const el = document.querySelector('[data-testid="app-shell"]') as HTMLElement;
      return window.getComputedStyle(el).gridTemplateRows;
    });
    // Layout rows (in order): titlebar, toolbar, banner (auto-collapses
    // to 0 when no banner present), formulabar, gridrow, mobilebar,
    // sheettabs. Track index [1] is the toolbar.
    const toolbarTrackAtPhone = parseInt(rowsAtPhone.split(/\s+/)[1], 10);
    expect(toolbarTrackAtPhone).toBeGreaterThan(0);
    expect(toolbarTrackAtPhone).toBeLessThanOrEqual(48);

    // Laptop: toolbar's grid track is the full --toolbar-h (≥36px).
    await page.setViewportSize({ width: 1280, height: 800 });
    await page.goto('/');
    await waitForUniver(page);
    const rowsAtLaptop = await page.evaluate(() => {
      const el = document.querySelector('[data-testid="app-shell"]') as HTMLElement;
      return window.getComputedStyle(el).gridTemplateRows;
    });
    const toolbarTrackAtLaptop = parseInt(rowsAtLaptop.split(/\s+/)[1], 10);
    expect(toolbarTrackAtLaptop).toBeGreaterThanOrEqual(36);
  });
});
