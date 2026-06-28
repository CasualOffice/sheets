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

import { expect, test, chromium, type Browser } from '@playwright/test';
import { spawn, type ChildProcess } from 'node:child_process';
import { waitForUniver } from './_helpers';

/**
 * End-to-end of the full "share for co-editing" flow:
 *
 *   1. Owner opens the app, clicks File → Share for co-editing…
 *   2. Sets a password, hits Create — UI flips to the "Room ready" stage
 *      with both share URLs visible.
 *   3. Owner clicks "Open the room" → navigates to /r/<id>, sees the
 *      avatar stack with their own name.
 *   4. Joiner browser opens the same URL → blocked by password prompt →
 *      submits correct password → joins → sees both avatars on each
 *      side; selection changes propagate as remote cursors.
 *   5. Joiner downloads a copy via File → Download a copy (.xlsx).
 *   6. Joiner clicks Leave room → URL drops back to /, joiner avatar
 *      disappears from owner's stack.
 *
 * Setup mirrors coedit.spec.ts — we spin up our own Hocuspocus server on
 * a separate port so the test is hermetic.
 */

test.describe.configure({ mode: 'serial', retries: 3 });

// Heavy flow: 2 browser contexts + a child Hocuspocus process +
// xlsx export + upload. The default 30 s test timeout is tight on
// shared CI runners — give it 90 s so worker fork-out + upload
// don't push past the wall under parallel load.
test.setTimeout(90_000);

const SERVER_PORT = 3056;
const WS_URL = `ws://127.0.0.1:${SERVER_PORT}/yjs`;
const HTTP_URL = `http://127.0.0.1:${SERVER_PORT}`;

let serverProc: ChildProcess | null = null;
let browser: Browser | null = null;

test.beforeAll(async () => {
  serverProc = spawn('pnpm', ['--filter', '@casualoffice/collab', 'exec', 'tsx', 'src/index.ts'], {
    env: { ...process.env, PORT: String(SERVER_PORT), HOST: '127.0.0.1' },
    stdio: ['ignore', 'pipe', 'pipe'],
  });

  await new Promise<void>((resolveReady, reject) => {
    const timer = setTimeout(() => reject(new Error('server boot timed out')), 15_000);
    serverProc!.stdout?.on('data', (chunk: Buffer) => {
      if (chunk.toString().includes('websocket sync on')) {
        clearTimeout(timer);
        resolveReady();
      }
    });
    serverProc!.on('exit', (code) => {
      clearTimeout(timer);
      reject(new Error(`server exited early (code ${code})`));
    });
  });

  browser = await chromium.launch();
});

test.afterAll(async () => {
  await browser?.close();
  if (serverProc && !serverProc.killed) {
    serverProc.kill('SIGTERM');
    await new Promise((r) => setTimeout(r, 200));
  }
});

/**
 * Returns an init-script that runs in the page context BEFORE any of our
 * app code. It does three things:
 *   - Points the collab driver at our standalone WS port.
 *   - Reroutes `/api/*` fetches to the standalone server (the Vite dev
 *     origin doesn't ship our Fastify endpoints).
 *   - Pre-seeds the display name + "already prompted" flag so the name
 *     prompt doesn't fire mid-test.
 *
 * The init-script string is built with the params interpolated — easier
 * than wrestling with Playwright's tuple-arg generics.
 */
function installCollabEnv(displayName: string): string {
  const wsJson = JSON.stringify(WS_URL);
  const httpJson = JSON.stringify(HTTP_URL);
  const nameJson = JSON.stringify(displayName);
  return `
    (function () {
      window.__COLLAB_WS_URL__ = ${wsJson};
      var origFetch = window.fetch.bind(window);
      window.fetch = function (url, init) {
        var u = typeof url === 'string' ? url : url.toString();
        if (u.indexOf('/api/') === 0) return origFetch(${httpJson} + u, init);
        return origFetch(url, init);
      };
      try {
        localStorage.setItem('casual.collab.displayName', ${nameJson});
        localStorage.setItem('casual.collab.namePrompted', '1');
      } catch (_) { /* private mode */ }
    })();
  `;
}

test('share → password join → presence → download → leave', async ({ baseURL }) => {
  // ── Owner ──────────────────────────────────────────────────────────
  const ownerCtx = await browser!.newContext();
  const owner = await ownerCtx.newPage();
  await owner.addInitScript({ content: installCollabEnv('Alice') });
  await owner.goto(baseURL!);
  await waitForUniver(owner);

  // Open the share dialog from the File menu.
  await owner.getByTestId('menubar-file').click();
  await owner.getByTestId('menu-item-start-room').click();
  await expect(owner.getByTestId('share-room-dialog')).toBeVisible();

  // Set a password and create the room.
  const PASSWORD = 'pa55phrase';
  await owner.getByTestId('share-room-password').fill(PASSWORD);
  await owner.getByTestId('share-room-create').click();

  // "Ready" stage exposes both URLs — read the edit link.
  // 20s is plenty even on slow CI: the dialog does a worker xlsx
  // export + a fetch upload behind the click, both of which can take
  // ~1–2 s on shared runners.
  const writeUrlInput = owner.getByTestId('share-room-write-url');
  await expect(writeUrlInput).toBeVisible({ timeout: 20_000 });
  const writeUrl = await writeUrlInput.inputValue();
  expect(writeUrl).toMatch(/\/r\/[\w-]{6,}$/);
  const roomId = writeUrl.split('/r/')[1];
  expect(roomId.length).toBeGreaterThanOrEqual(6);

  // Owner enters the room.
  await owner.getByTestId('share-room-open').click();
  await owner.waitForURL(`**/r/${roomId}`);
  await waitForUniver(owner);

  // Owner should see their own avatar in the stack.
  await expect(owner.getByTestId('presence-avatars')).toBeVisible({ timeout: 10_000 });
  await expect(owner.getByTestId('presence-avatar').filter({ hasText: 'AL' })).toBeVisible();

  // ── Joiner ─────────────────────────────────────────────────────────
  const joinerCtx = await browser!.newContext();
  const joiner = await joinerCtx.newPage();
  await joiner.addInitScript({ content: installCollabEnv('Bob') });
  await joiner.goto(`${baseURL}/r/${roomId}`);
  await waitForUniver(joiner);

  // Password gate fires from /info pre-flight.
  await expect(joiner.getByTestId('collab-password-dialog')).toBeVisible({ timeout: 10_000 });
  await joiner.getByTestId('collab-password-input').fill(PASSWORD);
  await joiner.getByTestId('collab-password-submit').click();

  // Joiner now in the room; both peers see two avatars (self + the other).
  await expect(joiner.getByTestId('presence-avatars')).toBeVisible({ timeout: 10_000 });
  await expect(joiner.getByTestId('presence-avatar')).toHaveCount(2, { timeout: 10_000 });
  await expect(owner.getByTestId('presence-avatar')).toHaveCount(2, { timeout: 10_000 });

  // Move the joiner's selection — a remote cursor should appear on the
  // owner's grid.
  await joiner.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('D4').activate();
  });
  await expect(owner.getByTestId('presence-cursor')).toBeVisible({ timeout: 10_000 });
  await expect(owner.getByTestId('presence-cursor').locator('.presence-cursor__label')).toHaveText(
    'Bob',
  );

  // ── Joiner downloads a copy ────────────────────────────────────────
  await joiner.getByTestId('menubar-file').click();
  const downloadPromise = joiner.waitForEvent('download');
  await joiner.getByTestId('menu-item-download-room').click();
  const download = await downloadPromise;
  expect(await download.path()).toBeTruthy();

  // ── Joiner leaves ──────────────────────────────────────────────────
  await joiner.getByTestId('menubar-file').click();
  await joiner.getByTestId('menu-item-leave-room').click();
  await joiner.waitForURL((u) => !u.pathname.startsWith('/r/'));
  await waitForUniver(joiner);

  // Owner's avatar stack drops back to just themselves.
  await expect(owner.getByTestId('presence-avatar')).toHaveCount(1, { timeout: 15_000 });

  await ownerCtx.close();
  await joinerCtx.close();
});

test('view-only role disables outbound edits', async ({ baseURL }) => {
  const ownerCtx = await browser!.newContext();
  const owner = await ownerCtx.newPage();
  await owner.addInitScript({ content: installCollabEnv('Owner') });
  await owner.goto(baseURL!);
  await waitForUniver(owner);

  // Create an open (no-password) room via the API directly — saves UI work.
  const roomId = await owner.evaluate(async () => {
    const res = await fetch('/api/rooms', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: '{}',
    });
    return ((await res.json()) as { roomId: string }).roomId;
  });

  await owner.goto(`${baseURL}/r/${roomId}`);
  await waitForUniver(owner);
  await expect(owner.getByTestId('presence-avatars')).toBeVisible({ timeout: 10_000 });

  // View-role joiner.
  const viewerCtx = await browser!.newContext();
  const viewer = await viewerCtx.newPage();
  await viewer.addInitScript({ content: installCollabEnv('Reader') });
  await viewer.goto(`${baseURL}/r/${roomId}?role=view`);
  await waitForUniver(viewer);

  // View-only banner appears once the room is live.
  await expect(viewer.getByTestId('view-only-banner')).toBeVisible({ timeout: 10_000 });

  // Viewer types into A1 locally — owner must NOT see it (view role
  // suppresses outbound op-log writes in bridge.ts).
  await viewer.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'should-not-sync' });
  });

  // Wait a bit longer than awareness/observer cycles, then read on owner.
  await owner.waitForTimeout(1500);
  const val = await owner.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return ws.getRange('A1').getValue();
  });
  expect(val).toBeFalsy();

  // Owner's edit DOES reach the viewer (they're still a receiver).
  await owner.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('B2').setValue({ v: 'from-owner' });
  });
  await viewer.waitForFunction(
    () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange('B2').getValue() === 'from-owner';
    },
    null,
    { timeout: 5_000 },
  );

  await ownerCtx.close();
  await viewerCtx.close();
});
