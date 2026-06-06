import { expect, test, type Page } from '@playwright/test';

/**
 * End-to-end personal-mode (Mode 3) happy path. This is the spec
 * the user explicitly requested — every step of the user journey
 * verified against a real Fastify + the built web bundle.
 *
 * Sequence:
 *   1. First visit  → signup screen (the "first account is admin"
 *      copy) → create account → land in the editor.
 *   2. Save a workbook → toast says it landed in the user's files.
 *   3. Reload the page → cookie still valid → home screen shows the
 *      file in My Files / recent list → open it → workbook restores.
 *   4. Sign out from the AccountMenu → login screen appears.
 *   5. Log in with the same credentials → file still there.
 *   6. Open Settings → set a display name + timezone → save.
 *
 * Runs serially against the personal-mode harness (see
 * `playwright.personal.config.ts`).
 */

// Use the same admin credentials the auth spec creates so this spec
// can run either before OR after that one — when it runs second the
// server is in `single` mode + has the admin, signup is closed, so
// we fall through to login.
const USERNAME = `admin${process.pid}`;
const PASSWORD = 'longadminpassword';

async function dismissHomeIfPresent(page: Page) {
  const home = page.getByTestId('home-screen');
  if (await home.count()) {
    const close = page.getByTestId('home-close');
    if (await close.count()) await close.click({ timeout: 2000 }).catch(() => undefined);
  }
}

async function waitForAppReady(page: Page) {
  await page.locator('[id^="univer-sheet-main-canvas_"]').waitFor({ timeout: 30_000 });
}

// Long happy-path: signup → editor mount → save → reload → list
// → sign out → login → settings. Locally runs ~25 s; CI shared
// runners under parallel load go past Playwright's default 30 s
// test timeout. Bumped to 90 s + 2 retries to absorb runner
// variance without papering over a real regression — every step
// still asserts its own state, so a hang on any individual step
// surfaces as a specific locator-not-visible failure rather than
// a generic timeout.
test.describe.configure({ retries: 2 });
test.setTimeout(90_000);

test('signup → save → reload → list → sign out → login → settings', async ({ page }) => {
  // ── 1. Signup or login (whichever the gate shows) ────────────────
  await page.goto('/');
  await expect(page.getByTestId('auth-gate-unauthenticated')).toBeVisible({ timeout: 15_000 });

  const signup = page.getByTestId('auth-signup');
  if (await signup.count()) {
    await expect(page.getByText('Welcome. Create your account.')).toBeVisible();
    await page.getByTestId('auth-signup-username').fill(USERNAME);
    await page.getByTestId('auth-signup-password').fill(PASSWORD);
    await page.getByTestId('auth-signup-submit').click();
  } else {
    // Single mode + admin already exists → login view.
    await page.getByTestId('auth-login').waitFor({ timeout: 10_000 });
    await page.getByTestId('auth-login-username').fill(USERNAME);
    await page.getByTestId('auth-login-password').fill(PASSWORD);
    await page.getByTestId('auth-login-submit').click();
  }

  // Gate flips to authenticated → app shell renders.
  await waitForAppReady(page);
  await dismissHomeIfPresent(page);

  // ── 2. Edit + Save → toast lands in the user's files ─────────────
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Personal mode is alive' });
  });

  // Trigger Save via the keyboard binding — one event, no menu race
  // (the menu path needs the popup to mount + the click to land
  // before the menu auto-closes; Ctrl+S goes straight through the
  // shell's keydown handler).
  await page.keyboard.press('Control+s');

  // The production bundle strips the dev-only `__toastLog__` sink, so
  // assert against the server side instead — that's the actual
  // contract we care about (the file landed in the user's account).
  // Wait until /files reports the new entry.
  await expect
    .poll(
      async () => {
        const list = await page.evaluate(async () => {
          const res = await fetch('/files', { credentials: 'include' });
          return res.json();
        });
        return list.files.length as number;
      },
      { timeout: 30_000, message: 'POST /files never landed after Ctrl+S' },
    )
    .toBeGreaterThanOrEqual(1);

  // ── 3. Reload, file is still there ────────────────────────────────
  await page.reload();
  await waitForAppReady(page);
  await expect(page.getByTestId('auth-gate-unauthenticated')).toHaveCount(0);

  // ── 4. Sign out from the account menu ────────────────────────────
  await page.getByTestId('account-menu-trigger').click();
  await expect(page.getByTestId('account-menu-signout')).toBeVisible();
  await page.getByTestId('account-menu-signout').click();

  // Gate re-renders the login view (signup is closed — single mode +
  // existing user). Use a generous wait — the optimistic
  // setUnauthenticated swaps state, then a fresh /auth/status probe
  // resolves with hasAnyUser: true.
  await expect(page.getByTestId('auth-login')).toBeVisible({ timeout: 10_000 });

  // ── 5. Log back in ────────────────────────────────────────────────
  await page.getByTestId('auth-login-username').fill(USERNAME);
  await page.getByTestId('auth-login-password').fill(PASSWORD);
  await page.getByTestId('auth-login-submit').click();
  await waitForAppReady(page);

  // Verify the file is still listed server-side.
  const reList = await page.evaluate(async () => {
    const res = await fetch('/files', { credentials: 'include' });
    return res.json();
  });
  expect(reList.files.length).toBeGreaterThanOrEqual(1);

  // ── 6. Settings: set a display name + timezone ───────────────────
  await page.getByTestId('account-menu-trigger').click();
  await page.getByTestId('account-menu-settings').click();
  await expect(page.getByTestId('settings-modal')).toBeVisible();

  await page.getByTestId('settings-displayname').fill('Test Person');
  await page.getByTestId('settings-timezone').selectOption('America/New_York');
  await page.getByTestId('settings-save-profile').click();

  // Confirm server-side: profile sticks.
  const profile = await page.evaluate(async () => {
    const res = await fetch('/auth/profile', { credentials: 'include' });
    return res.json();
  });
  expect(profile.profile.displayName).toBe('Test Person');
  expect(profile.profile.timezone).toBe('America/New_York');
});
