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

import { expect, test, type Page } from '@playwright/test';

/**
 * Secure share-link PASSWORD layer e2e (sharing-model §6.1, client side).
 * Drives the password input added to `CreateRoomDialog`'s secure-link
 * section: mint a password-protected link, confirm it lands in the list
 * with the lock / "password" badge (read off `hasPassword` from
 * GET /files/:id/shares), and cross-check the public /meta endpoint
 * reports `hasPassword: true` for that token.
 *
 * Runs against the personal-mode harness (`playwright.personal.config.ts`)
 * — a real Fastify with `/files/:id/shares*` + the public
 * `/files/shares/link/:token/meta` mounted, plus `/api/rooms`.
 *
 * NOTE on JOIN coverage: the full join-with-password handshake (a
 * `?share=<token>` joiner being prompted for `?sp=` and the server
 * verifying it on the WS upgrade) needs the collab WS + the Univer fork
 * build and a second client — not attempted here. The server's
 * `resolveJoinRole` unit matrix on main already covers the
 * password-required / password-mismatch / correct-password branches,
 * and `parseShareMeta` + the `/meta` endpoint are unit-tested. This spec
 * verifies the MINT side + the password indicator on the link row.
 *
 * Run with:
 *   pnpm exec playwright test -c playwright.personal.config.ts share-link-password
 */

const USERNAME = 'casualadmin';
const PASSWORD = 'longadminpassword';

async function dismissHomeIfPresent(page: Page) {
  const home = page.getByTestId('home-screen');
  if (await home.count()) {
    const close = page.getByTestId('home-close');
    if (await close.count()) await close.click({ timeout: 2000 }).catch(() => undefined);
  }
}

async function waitForSignedInLanding(page: Page) {
  await Promise.any([
    page.getByTestId('home-files-grid').waitFor({ timeout: 30_000 }),
    page.getByTestId('home-empty').waitFor({ timeout: 30_000 }),
    page.locator('[id^="univer-sheet-main-canvas_"]').waitFor({ timeout: 30_000 }),
  ]);
}

async function openEditorCanvas(page: Page) {
  const canvas = page.locator('[id^="univer-sheet-main-canvas_"]');
  if (await canvas.count()) return;
  const firstFile = page.locator('[data-testid^="home-file-row-"]').first();
  if (await firstFile.count()) {
    await firstFile.click();
  } else {
    await page.getByTestId('home-new-blank').click();
  }
  await canvas.waitFor({ timeout: 30_000 });
}

async function signInOrUp(page: Page) {
  await page.goto('/');
  await expect(page.getByTestId('auth-gate-unauthenticated')).toBeVisible({ timeout: 15_000 });
  const signup = page.getByTestId('auth-signup');
  if (await signup.count()) {
    await page.getByTestId('auth-signup-username').fill(USERNAME);
    await page.getByTestId('auth-signup-password').fill(PASSWORD);
    await expect(page.getByTestId('auth-signup-username')).toHaveValue(USERNAME);
    await expect(page.getByTestId('auth-signup-password')).toHaveValue(PASSWORD);
    await page.getByTestId('auth-signup-submit').click();
  } else {
    await page.getByTestId('auth-login').waitFor({ timeout: 10_000 });
    await page.getByTestId('auth-login-username').fill(USERNAME);
    await page.getByTestId('auth-login-password').fill(PASSWORD);
    await expect(page.getByTestId('auth-login-username')).toHaveValue(USERNAME);
    await expect(page.getByTestId('auth-login-password')).toHaveValue(PASSWORD);
    await page.getByTestId('auth-login-submit').click();
  }
  await waitForSignedInLanding(page);
}

test.describe.configure({ retries: 2 });
test.setTimeout(120_000);

test('mint a password-protected secure link → row shows the password badge', async ({ page }) => {
  // ── 1. Sign in + open the editor ──────────────────────────────────
  await signInOrUp(page);
  await openEditorCanvas(page);
  await dismissHomeIfPresent(page);

  // ── 2. Edit + Save so the workbook materialises a server file ──────
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'share-link password spec' });
  });
  await page.keyboard.press('Control+s');

  await expect
    .poll(
      async () =>
        page.evaluate(async () => {
          const res = await fetch('/files', { credentials: 'include' });
          const body = await res.json();
          return body.files.length as number;
        }),
      { timeout: 30_000, message: 'POST /files never landed after Ctrl+S' },
    )
    .toBeGreaterThanOrEqual(1);

  // Reload + reopen from the home list so the serverFileId binding (the
  // gate for the secure-link section) is explicit.
  await page.reload();
  await waitForSignedInLanding(page);
  const firstFile = page.locator('[data-testid^="home-file-row-"]').first();
  if (await firstFile.count()) await firstFile.click();
  await page.locator('[id^="univer-sheet-main-canvas_"]').waitFor({ timeout: 30_000 });
  await dismissHomeIfPresent(page);

  // ── 3. Open the share dialog → Ready stage → secure-link section ───
  await page.getByTestId('titlebar-share').click();
  await expect(page.getByTestId('share-room-dialog')).toBeVisible();
  await page.getByTestId('share-room-create').click();
  const section = page.getByTestId('share-link-section');
  await expect(section).toBeVisible({ timeout: 30_000 });

  // ── 4. Mint an EDIT link WITH a password ──────────────────────────
  await page.getByTestId('share-link-role').selectOption('edit');
  const sharePw = 'hunter2-share';
  await page.getByTestId('share-link-password').fill(sharePw);
  await page.getByTestId('share-link-create').click();

  // The row lands with the password badge (driven by hasPassword from
  // GET /files/:id/shares — the server never returns the hash).
  const list = page.getByTestId('share-link-list');
  const item = list.getByTestId('share-link-item').first();
  await expect(item).toBeVisible({ timeout: 15_000 });
  await expect(item).toHaveAttribute('data-role', 'edit');
  await expect(item.getByTestId('share-link-password-badge')).toBeVisible();

  // Password field is cleared after a successful mint (so a second link
  // isn't accidentally minted with the same stale value).
  await expect(page.getByTestId('share-link-password')).toHaveValue('');

  // ── 5. Cross-check the PUBLIC /meta endpoint reports hasPassword ──
  const url = await item.getByTestId('share-link-url').inputValue();
  const token = url.match(/[?&]share=([^&]+)/)?.[1];
  expect(token).toBeTruthy();
  const meta = await page.evaluate(async (t) => {
    const res = await fetch(`/files/shares/link/${t}/meta`);
    return res.json();
  }, token);
  expect(meta.valid).toBe(true);
  expect(meta.hasPassword).toBe(true);
  expect(meta.role).toBe('edit');
  // The bcrypt hash must NEVER appear in the public /meta response.
  expect(meta.passwordHash).toBeUndefined();
});
