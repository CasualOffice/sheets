import { defineConfig, devices } from '@playwright/test';

/**
 * Phase D / Mode 2 (WOPI embedded host) e2e harness.
 *
 * Boots a Fastify with:
 *   - `CASUAL_JWT_SECRET` set (enables token issuance + verification)
 *   - `CASUAL_ADMIN_USERNAME/PASSWORD` (the admin login that mints
 *     the test token via `POST /api/tokens`)
 *   - `CASUAL_STORAGE=local` + a scratch workbook dir on /tmp
 *
 * Tests run against http://127.0.0.1:3066. They sign in as admin
 * over the admin /admin endpoint to mint an access_token bound to
 * a seeded file, then visit `/?access_token=…` to drive the
 * `WopiFileSource` end-to-end.
 *
 * Default playwright config excludes `tests/e2e/wopi/**`; this
 * config opts them in.
 */

const SCRATCH = `/tmp/casual-wopi-e2e-${process.pid}`;
const PORT = 3066;

export default defineConfig({
  testDir: './tests/e2e/wopi',
  fullyParallel: false,
  workers: 1,
  reporter: process.env.CI ? [['github'], ['list']] : 'list',
  retries: process.env.CI ? 2 : 0,
  use: {
    baseURL: `http://127.0.0.1:${PORT}`,
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
    command: [
      `mkdir -p ${SCRATCH}/workbooks`,
      `pnpm --filter @sheet/web build`,
      `CASUAL_JWT_SECRET=wopi-e2e-secret-min-16chars-fine CASUAL_ADMIN_USERNAME=admin CASUAL_ADMIN_PASSWORD=adminpassword CASUAL_STORAGE=local CASUAL_LOCAL_PATH=${SCRATCH}/workbooks PORT=${PORT} RATE_LIMIT_ENABLED=false pnpm --filter @casualoffice/collab start`,
    ].join(' && '),
    url: `http://127.0.0.1:${PORT}/health`,
    timeout: 180_000,
    reuseExistingServer: false,
    stdout: 'pipe',
    stderr: 'pipe',
  },
});
