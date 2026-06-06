import { defineConfig, devices } from '@playwright/test';

/**
 * Playwright config for the personal-mode (Mode 3) e2e suite — the
 * tests that need a Fastify server running with
 * `CASUAL_PERSONAL_MODE=single|multi`.
 *
 * The default `playwright.config.ts` boots Vite dev, which doesn't
 * have a server attached and treats `/auth/status` as the SPA
 * fallback → the auth gate stays in `disabled` mode and these
 * specs would be no-ops.
 *
 * This config:
 *   1. Builds the web bundle once (so Fastify can static-serve it).
 *   2. Boots Fastify on port 3055 with the personal-mode env in
 *      place + a fresh users.db under a /tmp scratch dir.
 *   3. Targets `tests/e2e/personal/`.
 *
 * Run with:
 *   pnpm exec playwright test -c playwright.personal.config.ts
 *
 * CI: a dedicated job (added in the same Batch 5 push) builds the
 * web bundle, then runs this config.
 */

const SCRATCH = `/tmp/casual-personal-e2e-${process.pid}`;
const PORT = 3055;

export default defineConfig({
  testDir: './tests/e2e/personal',
  // These specs share a single Fastify server, so run them
  // serially — parallel-running tests would race on the same
  // users.db.
  fullyParallel: false,
  workers: 1,
  reporter: process.env.CI ? [['github'], ['list']] : 'list',
  retries: process.env.CI ? 2 : 0,
  use: {
    baseURL: `http://127.0.0.1:${PORT}`,
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
    video: 'retain-on-failure',
    // Personal mode flows always involve a cookie; keep it.
    storageState: undefined,
  },
  projects: [
    {
      name: 'chromium',
      use: { ...devices['Desktop Chrome'] },
    },
  ],
  webServer: {
    // 1) Build the web bundle if it's not there. 2) Spin up Fastify
    // in `single` mode, pointed at a scratch dir on /tmp so the
    // operator's real /data isn't touched.
    command: [
      `mkdir -p ${SCRATCH}/workbooks`,
      `rm -f ${SCRATCH}/users.db ${SCRATCH}/users.db-wal ${SCRATCH}/users.db-shm`,
      `pnpm --filter @sheet/web build`,
      `CASUAL_PERSONAL_MODE=single CASUAL_USERS_DB_PATH=${SCRATCH}/users.db CASUAL_STORAGE=local CASUAL_LOCAL_PATH=${SCRATCH}/workbooks PORT=${PORT} RATE_LIMIT_ENABLED=false pnpm --filter @sheet/server start`,
    ].join(' && '),
    url: `http://127.0.0.1:${PORT}/health`,
    timeout: 180_000,
    reuseExistingServer: false,
    stdout: 'pipe',
    stderr: 'pipe',
  },
});
