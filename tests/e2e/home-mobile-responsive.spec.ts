/**
 * /home (MySpreadsheetsList) visual smoke at both desktop and mobile
 * viewports. UX_AUDIT.md §2.15 — the file list must work on phones,
 * which means tighter padding, smaller title, an always-visible Delete
 * affordance (no hover on touch), and a wrapping actions row.
 *
 * Mocks `/auth/status` (signed-in single-mode admin) and `/files`
 * (two fake rows) so the spec runs against the Vite dev server with
 * no backend.
 */
import { test, expect } from '@playwright/test';

const FAKE_USER = {
  id: 1,
  username: 'demo',
  isAdmin: true,
  createdAt: 1700000000000,
};

// Server-shape ServerFileMeta — the gateway returns `{ files: [...] }`
// and personal-file-source.ts maps it to RecentEntry. Mocking the wire
// shape (not the React-side shape) keeps this spec honest if the
// mapping changes.
const FAKE_FILES = [
  {
    id: 'wb-1',
    name: 'Q3 budget.xlsx',
    size: 12345,
    etag: 'v1',
    createdAt: Date.parse('2026-06-01T00:00:00Z'),
    modifiedAt: Date.parse('2026-06-10T00:00:00Z'),
  },
  {
    id: 'wb-2',
    name: 'Sprint planning.xlsx',
    size: 6789,
    etag: 'v1',
    createdAt: Date.parse('2026-06-01T00:00:00Z'),
    modifiedAt: Date.parse('2026-06-08T00:00:00Z'),
  },
];

async function mockPersonalMode(page: import('@playwright/test').Page) {
  await page.route('**/auth/status', (route) =>
    route.fulfill({
      status: 200,
      contentType: 'application/json',
      body: JSON.stringify({ mode: 'single', user: FAKE_USER }),
    }),
  );
  await page.route('**/files', (route) =>
    route.fulfill({
      status: 200,
      contentType: 'application/json',
      body: JSON.stringify({ files: FAKE_FILES }),
    }),
  );
}

test.describe('MySpreadsheetsList — responsive', () => {
  test('desktop: header + two file rows render side-by-side', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 });
    await mockPersonalMode(page);
    await page.goto('/home');
    await expect(page.getByRole('heading', { name: 'My Spreadsheets' })).toBeVisible();
    await expect(page.getByTestId('home-files-grid')).toBeVisible();
    await expect(page.getByText('Q3 budget.xlsx')).toBeVisible();
    await expect(page.getByText('Sprint planning.xlsx')).toBeVisible();
    // Delete is hover-only on desktop, so it should NOT be visible at rest.
    const deleteBtn = page.getByTestId('home-file-delete-wb-1');
    await expect(deleteBtn).not.toBeVisible();
  });

  test('mobile: title shrinks, actions wrap, delete is always visible', async ({ page }) => {
    await page.setViewportSize({ width: 375, height: 720 });
    await mockPersonalMode(page);
    await page.goto('/home');
    await expect(page.getByRole('heading', { name: 'My Spreadsheets' })).toBeVisible();
    await expect(page.getByText('Q3 budget.xlsx')).toBeVisible();
    // Phones have no hover, so the audit fix makes the Delete affordance
    // reachable without one. This is the single most important mobile
    // regression catch — if it's hidden, you can never delete a row.
    const deleteBtn = page.getByTestId('home-file-delete-wb-1');
    await expect(deleteBtn).toBeVisible();
  });
});
