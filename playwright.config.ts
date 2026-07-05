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
  // CI shards across 4 runners; cap workers to the 2 vCPUs per runner so we
  // don't oversubscribe the CPU and slow every test into timeout territory
  // (the chart/format specs were flaking on assertion timeouts under load).
  workers: process.env.CI ? 2 : undefined,
  // Give assertions headroom under CI contention.
  expect: { timeout: process.env.CI ? 10000 : 5000 },
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
    '**/coedit-comments.spec.ts',
    '**/coedit-drawings.spec.ts', // runs against prod stack via playwright.docker.config.ts
    // coedit-compaction spawns its own Hocuspocus on :3060 and drives a heavy
    // 220-mutation → op-log-compaction → joiner-replay flow. Against the cold
    // Vite dev server under CI contention it times out even at 90s waits (it was
    // the lone chronic e2e failure on main). Excluded here like its siblings;
    // re-home it to the prod-stack docker job (playwright.docker.config.ts) where
    // the fast built bundle makes the timing reliable.
    '**/coedit-compaction.spec.ts',
    // Personal-mode (Mode 3) specs need a Fastify with
    // CASUAL_PERSONAL_MODE=single; they run via
    // playwright.personal.config.ts which manages its own server.
    'tests/e2e/personal/**',
    // WOPI (Mode 2 / Phase D) specs need a Fastify with
    // CASUAL_JWT_SECRET + an admin to mint tokens; run via
    // playwright.wopi.config.ts.
    'tests/e2e/wopi/**',
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
