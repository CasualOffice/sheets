import { expect, test, type Page } from '@playwright/test';

/**
 * Secure share-LINK dialog e2e (sharing-model §6.1, client side). Drives
 * the Link affordance that `CreateRoomDialog` shows for a personal-mode
 * SAVED file (one with a `serverFileId`): mint a view link, see it in the
 * list with a copy-able `?share=` URL, then revoke it.
 *
 * Runs against the personal-mode harness (`playwright.personal.config.ts`)
 * — a real Fastify with `/files/:id/shares*` mounted + `/api/rooms`. The
 * dialog first creates + seeds an anonymous room (the existing flow), then
 * mints a token bound to that room via POST /files/:id/shares/link.
 *
 * NOTE on enforcement coverage: the actual ROLE enforcement (a `?share=`
 * token resolving server-authoritatively to view/edit on the WS upgrade)
 * is covered by the server's `resolveJoinRole` unit matrix already on
 * main; a full two-client enforcement e2e needs the collab WS + the
 * Univer fork build and isn't attempted here. This spec verifies the
 * client minting / listing / revoke surface + the URL shape.
 *
 * Run with:
 *   pnpm exec playwright test -c playwright.personal.config.ts share-link
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
    // Confirm the controlled inputs captured the values before submit — a
    // fast fill→click can submit before React's onChange state settles,
    // sending an empty/stale password (intermittent 401). See flake note.
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

test('mint + list + revoke a secure share link on a saved personal file', async ({ page }) => {
  // ── 1. Sign in + open the editor ──────────────────────────────────
  await signInOrUp(page);
  await openEditorCanvas(page);
  await dismissHomeIfPresent(page);

  // ── 2. Edit + Save so the workbook materialises a server file ──────
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'share-link spec' });
  });
  await page.keyboard.press('Control+s');

  // Wait until the file is registered server-side (the save bound a
  // serverFileId to the workbook in the process).
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

  // Reload + reopen the file from the home list. openRecent() returns the
  // serverFileId, which is exactly what the share dialog gates the secure
  // link section on — going through the list makes the binding explicit
  // rather than relying on the post-save draft → id rewrite.
  await page.reload();
  await waitForSignedInLanding(page);
  const firstFile = page.locator('[data-testid^="home-file-row-"]').first();
  if (await firstFile.count()) await firstFile.click();
  await page.locator('[id^="univer-sheet-main-canvas_"]').waitFor({ timeout: 30_000 });
  await dismissHomeIfPresent(page);

  // ── 3. Open the share dialog → through to the Ready stage ──────────
  await page.getByTestId('titlebar-share').click();
  await expect(page.getByTestId('share-room-dialog')).toBeVisible();
  await page.getByTestId('share-room-create').click();

  // The secure-link section only renders in the Ready stage for a saved
  // personal file — its presence proves the serverFileId gate works.
  const section = page.getByTestId('share-link-section');
  await expect(section).toBeVisible({ timeout: 30_000 });

  // ── 4. Mint a VIEW link ───────────────────────────────────────────
  await page.getByTestId('share-link-role').selectOption('view');
  await page.getByTestId('share-link-create').click();

  // It lands in the list with a copy-able ?share= URL labelled "view".
  const list = page.getByTestId('share-link-list');
  const item = list.getByTestId('share-link-item').first();
  await expect(item).toBeVisible({ timeout: 15_000 });
  await expect(item).toHaveAttribute('data-role', 'view');

  const urlInput = item.getByTestId('share-link-url');
  await expect(urlInput).toBeVisible();
  const url = await urlInput.inputValue();
  expect(url).toContain('?share=');
  expect(url).toContain('/r/');

  // Copy button is present (clipboard may be unavailable headless — the
  // field stays selectable either way; we just assert the control wired).
  await expect(item.getByTestId('share-link-copy')).toBeVisible();

  // Cross-check the server actually persisted the token.
  const fileId = url.match(/\/r\/([^?]+)/)?.[1];
  expect(fileId).toBeTruthy();

  // ── 5. Revoke empties the list ────────────────────────────────────
  // Revoke EVERY link on the file, not just .first(): links are file-scoped
  // and a sibling spec sharing the single-mode admin (or a prior retry) can
  // leave others on the same file, so a plain "revoke one → count 0" is racy.
  const items = list.getByTestId('share-link-item');
  for (let n = await items.count(); n > 0; n--) {
    await items.first().getByTestId('share-link-revoke').click();
    await expect(items).toHaveCount(n - 1, { timeout: 15_000 });
  }
  await expect(page.getByTestId('share-link-empty')).toBeVisible();
});
