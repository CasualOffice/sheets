import { defineConfig, devices } from '@playwright/test';

/**
 * Industry-standard Playwright config:
 *  - Auto-starts the web dev server.
 *  - Chromium first; expand to webkit/firefox in Phase 3.
 *  - Trace + screenshot on first failure for actionable debugging.
 */
export default defineConfig({
  testDir: './tests/e2e',
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  reporter: process.env.CI ? [['github'], ['html', { open: 'never' }]] : 'list',
  // These specs hit the production docker stack on :3000 (Fastify
  // serves built bundle + api + WS from one origin) — they're run
  // separately via playwright.docker.config.ts in the e2e-prod CI
  // job. Excluding them here so the default dev-server e2e doesn't
  // try to connect to :3000 and fail with ERR_CONNECTION_REFUSED.
  testIgnore: [
    '**/coedit-cell-edit.spec.ts',
    '**/coedit-view-only.spec.ts',
    '**/coedit-divergence.spec.ts',
    '**/coedit-insert-sheet.spec.ts',
    '**/coedit-sheet-ops.spec.ts',
    '**/coedit-cursor-math.spec.ts',
    '**/coedit-synced-mutations.spec.ts',
    '**/coedit-conditional-formatting.spec.ts',
    '**/coedit-data-validation.spec.ts',
    '**/coedit-workbook-metadata.spec.ts',
    '**/coedit-drawings.spec.ts',
  ],
  use: {
    baseURL: 'http://127.0.0.1:5273',
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
    video: 'retain-on-failure',
  },
  projects: [
    {
      name: 'chromium',
      use: { ...devices['Desktop Chrome'] },
    },
  ],
  webServer: {
    command: 'pnpm --filter @sheet/web dev',
    url: 'http://127.0.0.1:5273',
    reuseExistingServer: !process.env.CI,
    stdout: 'pipe',
    stderr: 'pipe',
    timeout: 120_000,
  },
});
