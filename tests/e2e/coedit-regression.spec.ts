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
 * Regression suite for the May-2026 co-edit audit. Each test maps to a
 * row in docs/COLLAB-FIXES.md:
 *
 *   - "wrong password surfaces in the password prompt" → Issue 1
 *   - "remote sheet creation does not switch local active sheet" → Issue 20
 *
 * Shares the standalone-server harness pattern from coedit-share.spec.ts;
 * the helpers live there because they're already battle-tested for the
 * collab tests.
 */

test.describe.configure({ mode: 'serial' });

const SERVER_PORT = 3058;
const WS_URL = `ws://127.0.0.1:${SERVER_PORT}/yjs`;
const HTTP_URL = `http://127.0.0.1:${SERVER_PORT}`;

let serverProc: ChildProcess | null = null;
let browser: Browser | null = null;

test.beforeAll(async () => {
  serverProc = spawn(
    'pnpm',
    ['--filter', '@casualoffice/collab', 'exec', 'tsx', 'src/index.ts'],
    {
      env: { ...process.env, PORT: String(SERVER_PORT), HOST: '127.0.0.1' },
      stdio: ['ignore', 'pipe', 'pipe'],
    },
  );
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

test('wrong password rejects → user sees the password prompt with the error message', async ({ baseURL }) => {
  // Set up an owner who creates a password-protected room.
  const ownerCtx = await browser!.newContext();
  const owner = await ownerCtx.newPage();
  await owner.addInitScript({ content: installCollabEnv('Owner') });
  await owner.goto(baseURL!);
  await waitForUniver(owner);

  // Create the room via the API directly — simpler than going through
  // the dialog flow.
  const roomId = await owner.evaluate(async () => {
    const res = await fetch('/api/rooms', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ password: 'correct-horse-battery' }),
    });
    return ((await res.json()) as { roomId: string }).roomId;
  });

  // Joiner submits the WRONG password. Before the fix, the server's
  // pre-upgrade HTTP 401 surfaced as WebSocket close code 1006 — which
  // the client only matched on 401/4401, so the password prompt never
  // re-opened and the provider silently retried forever. The fix
  // accepts the upgrade then closes with code 4401 so the client sees
  // a real close code.
  const joinerCtx = await browser!.newContext();
  const joiner = await joinerCtx.newPage();
  await joiner.addInitScript({ content: installCollabEnv('Joiner') });
  await joiner.goto(`${baseURL}/r/${roomId}`);

  // First prompt fires from /info pre-flight.
  await expect(joiner.getByTestId('collab-password-dialog')).toBeVisible({ timeout: 10_000 });
  await joiner.getByTestId('collab-password-input').fill('wrong-passphrase');
  await joiner.getByTestId('collab-password-submit').click();

  // The dialog should re-open with the error message — this is the
  // assertion that fails without the fix (dialog stays closed forever,
  // status stuck at "connecting").
  await expect(joiner.getByTestId('collab-password-dialog')).toBeVisible({ timeout: 10_000 });
  await expect(joiner.getByTestId('collab-password-error')).toBeVisible({ timeout: 5_000 });
  await expect(joiner.getByTestId('collab-password-error')).toContainText(/incorrect|try again/i);

  // Submitting the correct password lets them in.
  await joiner.getByTestId('collab-password-input').fill('correct-horse-battery');
  await joiner.getByTestId('collab-password-submit').click();
  await expect(joiner.getByTestId('presence-avatars')).toBeVisible({ timeout: 10_000 });

  await ownerCtx.close();
  await joinerCtx.close();
});

test('remote insert-sheet does not switch local user away from their active sheet', async ({ baseURL }) => {
  // Owner creates an open room and stays on Sheet1.
  const ownerCtx = await browser!.newContext();
  const owner = await ownerCtx.newPage();
  await owner.addInitScript({ content: installCollabEnv('Alice') });
  await owner.goto(baseURL!);
  await waitForUniver(owner);

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

  // Capture owner's active sheet id BEFORE the joiner adds a sheet.
  const ownerSheetBefore = await owner.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (api.getActiveWorkbook()!.getActiveSheet() as any).getSheetId();
  });

  // Joiner connects.
  const joinerCtx = await browser!.newContext();
  const joiner = await joinerCtx.newPage();
  await joiner.addInitScript({ content: installCollabEnv('Bob') });
  await joiner.goto(`${baseURL}/r/${roomId}`);
  await waitForUniver(joiner);
  await expect(joiner.getByTestId('presence-avatars')).toBeVisible({ timeout: 10_000 });

  // Joiner inserts a new sheet. Univer's ActiveWorksheetController fires
  // _adjustActiveSheetOnInsertSheet on every insert-sheet mutation — even
  // ones replayed with fromCollab: true. Without the bridge's save+
  // restore, the owner's active sheet would silently switch to the new
  // one Bob just created.
  await joiner.evaluate(() => {
    const api = window.__univerAPI!;
    api.getActiveWorkbook()!.insertSheet();
  });

  // Give the bridge a beat to replay.
  await owner.waitForTimeout(1500);

  const ownerSheetAfter = await owner.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (api.getActiveWorkbook()!.getActiveSheet() as any).getSheetId();
  });
  expect(ownerSheetAfter).toBe(ownerSheetBefore);

  // But the new sheet IS present in the owner's tab strip (the mutation
  // itself propagated — only the active-sheet switch was suppressed).
  const ownerSheetCount = await owner.evaluate(() => {
    const api = window.__univerAPI!;
    return api.getActiveWorkbook()!.getSheets().length;
  });
  expect(ownerSheetCount).toBeGreaterThanOrEqual(2);

  await ownerCtx.close();
  await joinerCtx.close();
});
