import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

// Phone-sized viewport. Drives the @media (max-width: 480px) / 720px
// rules in styles.css.
const PHONE_VIEWPORT = { width: 375, height: 667 };

/**
 * Phone-viewport smoke tests. Univer's grid handles touch natively; the
 * job here is making sure our React shell doesn't break or hide critical
 * surfaces on a 375 px screen.
 *
 * Not a "full mobile editing" suite — editing on touch is a Phase 3
 * polish job. These tests just lock down that the app renders without
 * horizontal overflow and the essential nav surfaces are still reachable.
 */
test.describe('Mobile shell (375 × 667)', () => {
  test.beforeEach(async ({ page }) => {
    await page.setViewportSize(PHONE_VIEWPORT);
  });

  test('renders without horizontal overflow', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    const overflow = await page.evaluate(() => {
      const docW = document.documentElement.scrollWidth;
      const viewW = window.innerWidth;
      return { docW, viewW };
    });
    // Document width may slightly exceed viewport because the menubar +
    // toolbar are designed to horizontally scroll on phones. The BODY
    // itself shouldn't overflow.
    const bodyOverflow = await page.evaluate(() => {
      const body = document.body;
      return body.scrollWidth - body.clientWidth;
    });
    expect(bodyOverflow, `body overflowed by ${bodyOverflow}px (viewport ${overflow.viewW})`).toBeLessThanOrEqual(2);
  });

  test('menu bar is reachable and scrolls horizontally to reveal all items', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

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

  test('toolbar is hidden on a 375 px viewport but the formula bar stays', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    // Below 480 px the toolbar is intentionally hidden — menus carry the
    // same commands. Formula bar + grid are the working surface.
    await expect(page.getByTestId('toolbar')).toBeHidden();
    await expect(page.getByTestId('formula-bar')).toBeVisible();
    await expect(page.getByTestId('grid-host')).toBeVisible();
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
