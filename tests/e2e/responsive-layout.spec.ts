import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Polish #4 — responsive layout. Locks in the grid-host always
 * fills the available vertical space + sheet-tabs always sits at
 * the bottom + the toolbar collapses cleanly on the narrow phone
 * breakpoint.
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
    test(`grid-host fills the viewport and sheet-tabs sits at the bottom @ ${v.name} (${v.w}x${v.h})`, async ({
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
          return { w: Math.round(r.width), h: Math.round(r.height), top: Math.round(r.top), bottom: Math.round(r.bottom) };
        };
        return {
          app: q('[data-testid="app-shell"]'),
          gridHost: q('.grid-host'),
          sheetTabs: q('.sheet-tabs'),
          viewportH: window.innerHeight,
        };
      });

      expect(layout.app).not.toBeNull();
      expect(layout.gridHost).not.toBeNull();
      expect(layout.sheetTabs).not.toBeNull();

      // 1. App shell fills the full viewport (100dvh).
      expect(layout.app!.h).toBeGreaterThanOrEqual(layout.viewportH - 1);

      // 2. Grid host gets the bulk of the vertical space — at least
      //    60 % of the viewport at any size. Earlier bug had the host
      //    at 26 px (3% of an 800 px viewport).
      expect(layout.gridHost!.h).toBeGreaterThan(layout.viewportH * 0.6);

      // 3. Sheet tabs strip is at the bottom — its bottom edge sits
      //    within 2 px of the viewport bottom.
      expect(layout.sheetTabs!.bottom).toBeGreaterThanOrEqual(layout.viewportH - 2);
      expect(layout.sheetTabs!.bottom).toBeLessThanOrEqual(layout.viewportH + 1);
    });
  }

  test('phone breakpoint (<= 480) collapses the toolbar grid track', async ({ page }) => {
    await page.setViewportSize({ width: 480, height: 800 });
    await page.goto('/');
    await waitForUniver(page);
    const toolbar = await page.evaluate(() => {
      const el = document.querySelector('.toolbar') as HTMLElement | null;
      const cs = el ? window.getComputedStyle(el) : null;
      return { display: cs?.display, rect: el?.getBoundingClientRect() };
    });
    expect(toolbar.display).toBe('none');
    expect(toolbar.rect?.height ?? 0).toBe(0);
  });

  test('grid track for toolbar is 0px at phone, full --toolbar-h at wider viewports', async ({
    page,
  }) => {
    // Phone: toolbar's grid track collapses; the gridTemplateRows
    // third value should report `0px`.
    await page.setViewportSize({ width: 480, height: 800 });
    await page.goto('/');
    await waitForUniver(page);
    const rowsAtPhone = await page.evaluate(() => {
      const el = document.querySelector('[data-testid="app-shell"]') as HTMLElement;
      return window.getComputedStyle(el).gridTemplateRows;
    });
    // Format: "40px 28px 0px 26px Npx 36px" — track 3 (toolbar) is 0px.
    expect(rowsAtPhone.split(/\s+/)[2]).toBe('0px');

    // Laptop: toolbar's grid track is the full --toolbar-h.
    await page.setViewportSize({ width: 1280, height: 800 });
    await page.goto('/');
    await waitForUniver(page);
    const rowsAtLaptop = await page.evaluate(() => {
      const el = document.querySelector('[data-testid="app-shell"]') as HTMLElement;
      return window.getComputedStyle(el).gridTemplateRows;
    });
    const toolbarTrackAtLaptop = parseInt(rowsAtLaptop.split(/\s+/)[2], 10);
    expect(toolbarTrackAtLaptop).toBeGreaterThanOrEqual(36);
  });
});
