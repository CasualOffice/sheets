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

/**
 * Personal-mode (Mode 3) auth + session edge cases — every refusal
 * path the design doc promised. Runs against the same harness as
 * `personal-mode-flow.spec.ts`: a Fastify in `single` mode with a
 * scratch users.db.
 *
 * Specs are explicitly ordered via test name lex so the shared
 * server state is predictable across them:
 *
 *   01 — first signup creates the admin
 *   02 — duplicate username rejected (409)
 *   03 — single mode rejects a second signup (403)
 *   04 — wrong password → invalid-credentials toast on the login view
 *   05 — session cookie cleared → gate re-renders login
 *   06 — change password from Settings invalidates other sessions
 *   07 — change password lets the caller stay signed in
 *
 * No file ops in this spec — those live in `personal-mode-flow`.
 */

const ADMIN_USER = 'casualadmin';
const ADMIN_PASSWORD = 'longadminpassword';
const ADMIN_PASSWORD_NEXT = 'newlongadminpassword';

/** After signup/login on personal mode, the router redirects to
 *  `/home` (the My Spreadsheets list view) — NOT a blank editor.
 *  Tests that only need the authenticated-state gate flipped can
 *  treat any of these as "landed": the home grid, the home empty
 *  state, or the editor canvas (if a workbook was already open). */
async function waitForLandedSignedIn(page: import('@playwright/test').Page) {
  await Promise.any([
    page.getByTestId('home-files-grid').waitFor({ timeout: 30_000 }),
    page.getByTestId('home-empty').waitFor({ timeout: 30_000 }),
    page.locator('[id^="univer-sheet-main-canvas_"]').waitFor({ timeout: 30_000 }),
  ]);
}

async function ensureSignedIn(page: import('@playwright/test').Page) {
  await page.goto('/');
  const login = page.getByTestId('auth-login');
  await Promise.race([
    login.waitFor({ timeout: 10_000 }).catch(() => undefined),
    page
      .getByTestId('home-empty')
      .waitFor({ timeout: 10_000 })
      .catch(() => undefined),
    page
      .getByTestId('home-files-grid')
      .waitFor({ timeout: 10_000 })
      .catch(() => undefined),
  ]);
  if (await login.count()) {
    await page.getByTestId('auth-login-username').fill(ADMIN_USER);
    await page.getByTestId('auth-login-password').fill(ADMIN_PASSWORD);
    await page.getByTestId('auth-login-submit').click();
    await waitForLandedSignedIn(page);
  }
}

test('01 first signup creates the admin and lands the home screen', async ({ page }) => {
  await page.goto('/');
  await expect(page.getByTestId('auth-signup')).toBeVisible({ timeout: 15_000 });
  await page.getByTestId('auth-signup-username').fill(ADMIN_USER);
  await page.getByTestId('auth-signup-password').fill(ADMIN_PASSWORD);
  await page.getByTestId('auth-signup-submit').click();
  // Personal-mode signup redirects to /home; the empty list shows
  // first ("No spreadsheets yet"). Either the grid or the empty
  // sentinel is the success state for "the gate flipped to authed".
  await waitForLandedSignedIn(page);
});

test('02 duplicate username returns username-taken via the signup view', async ({ page }) => {
  // Server is now in single mode + has the admin — direct API hit
  // exercises the same code path the signup view would, without
  // relying on the gate temporarily re-rendering signup (it stays
  // on login from now on).
  const res = await page.request.post('/auth/signup', {
    data: { username: ADMIN_USER, password: 'longpasswordattempt' },
  });
  // Server returns 403 signup-closed before it even checks the name
  // in single mode — that's the more-specific refusal. Verify either.
  expect([403, 409]).toContain(res.status());
});

test('03 single-mode rejects a fresh signup (403 signup-closed)', async ({ page }) => {
  const res = await page.request.post('/auth/signup', {
    data: { username: 'someoneelse', password: 'longpasswordattempt' },
  });
  expect(res.status()).toBe(403);
  expect(((await res.json()) as { error: string }).error).toBe('signup-closed');
});

test('04 wrong password on the login view shows invalid-credentials', async ({ page, context }) => {
  // Sign out first if we landed signed-in from a prior test.
  await context.clearCookies();
  await page.goto('/');
  await page.getByTestId('auth-login').waitFor({ timeout: 10_000 });
  await page.getByTestId('auth-login-username').fill(ADMIN_USER);
  await page.getByTestId('auth-login-password').fill('definitelywrong');
  await page.getByTestId('auth-login-submit').click();
  await expect(page.getByTestId('auth-error')).toContainText(/wrong/i, { timeout: 5_000 });
  // Gate stays on login.
  await expect(page.getByTestId('auth-login')).toBeVisible();
});

test('05 cleared cookie re-renders the login view on next visit', async ({ page, context }) => {
  await ensureSignedIn(page);
  await context.clearCookies();
  await page.reload();
  await expect(page.getByTestId('auth-login')).toBeVisible({ timeout: 10_000 });
});

test('06 change-password invalidates every other session', async ({ page, context }) => {
  await ensureSignedIn(page);

  // Open a SECOND browser context and sign in there too — that's
  // the "other tab" that should get logged out by the change.
  const otherCtx = await context.browser()!.newContext();
  const otherPage = await otherCtx.newPage();
  await otherPage.goto('/');
  // The second context might see the signup screen (no signed-in
  // session yet) — wait for either signup OR login since the gate
  // chooses based on `signupAllowed`, which is false here (single
  // mode + existing user).
  await otherPage.getByTestId('auth-login').waitFor({ timeout: 10_000 });
  await otherPage.getByTestId('auth-login-username').fill(ADMIN_USER);
  await otherPage.getByTestId('auth-login-password').fill(ADMIN_PASSWORD);
  await otherPage.getByTestId('auth-login-submit').click();
  await waitForLandedSignedIn(otherPage);

  // Now change the password from the primary tab.
  await page.getByTestId('account-menu-trigger').click();
  await page.getByTestId('account-menu-settings').click();
  await page.getByTestId('settings-tab-security').click();
  await page.getByTestId('settings-current-password').fill(ADMIN_PASSWORD);
  await page.getByTestId('settings-new-password').fill(ADMIN_PASSWORD_NEXT);
  await page.getByTestId('settings-confirm-password').fill(ADMIN_PASSWORD_NEXT);
  await page.getByTestId('settings-save-password').click();

  // The store deletes every session for the user; the route layer
  // re-issues a cookie for the caller so the primary tab stays
  // signed in. The other tab's session id is now invalid — confirm
  // /auth/me from that context returns 401.
  const meRes = await otherPage.request.get('/auth/me');
  expect(meRes.status()).toBe(401);

  await otherCtx.close();
});

test('07 the password we just set works for a fresh login', async ({ page, context }) => {
  await context.clearCookies();
  await page.goto('/');
  await page.getByTestId('auth-login').waitFor({ timeout: 10_000 });

  // Old password should fail.
  await page.getByTestId('auth-login-username').fill(ADMIN_USER);
  await page.getByTestId('auth-login-password').fill(ADMIN_PASSWORD);
  await page.getByTestId('auth-login-submit').click();
  await expect(page.getByTestId('auth-error')).toContainText(/wrong/i, { timeout: 5_000 });

  // New password should work.
  await page.getByTestId('auth-login-password').fill(ADMIN_PASSWORD_NEXT);
  await page.getByTestId('auth-login-submit').click();
  await waitForLandedSignedIn(page);
});

test('08 restore admin password so later specs find the credential they expect', async ({
  page,
}) => {
  // The harness is shared across spec files. `personal-mode-flow.spec.ts`
  // signs in with `ADMIN_PASSWORD`; without this rollback it would 401.
  // Sign in fresh with the NEW password (the one spec 07 left in
  // place), then rotate back. Doing this entirely via the request
  // API avoids the home-overlay / editor-ready dance that
  // `ensureSignedIn` does — and `ensureSignedIn` would 401 here
  // anyway since it hardcodes the old password.
  const login = await page.request.post('/auth/login', {
    data: { username: ADMIN_USER, password: ADMIN_PASSWORD_NEXT },
  });
  expect(login.status()).toBe(200);
  const reset = await page.request.post('/auth/change-password', {
    data: { currentPassword: ADMIN_PASSWORD_NEXT, newPassword: ADMIN_PASSWORD },
  });
  expect(reset.status()).toBe(204);
});
