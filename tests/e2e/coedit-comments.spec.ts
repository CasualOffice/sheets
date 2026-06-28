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

import { expect, test, chromium, type Browser, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Real-time comment sync across co-editors.
 *
 * Thread comments (the `thread-comment.mutation.*` family) are emitted as
 * `CommandType.MUTATION`, so they ride the same op-log bridge as every
 * other synced mutation — BUT only if they're in `SYNCED_MUTATIONS`
 * (apps/web → packages/sdk/src/collab/bridge.ts) and the joiner loads the
 * `threadComment` lazy plugin group before replay (MUTATION_TO_LAZY_GROUP).
 * Before this slice they were in neither, so a comment added in a shared
 * room never reached co-editors. These tests fire the comment command on
 * peer A and assert the thread shows up on peer B — a regression guard for
 * that wiring.
 *
 * Mirrors coedit-synced-mutations.spec.ts for the two-client room setup.
 */

const PROD_BASE = process.env.PROD_BASE ?? 'http://localhost:3000';

let browser: Browser | null = null;
test.describe.configure({ mode: 'serial' });

test.beforeAll(async () => {
  browser = await chromium.launch();
});
test.afterAll(async () => {
  await browser?.close();
});

function installEnv(name: string): string {
  return `
    (function () {
      try {
        localStorage.setItem('casual.collab.displayName', ${JSON.stringify(name)});
        localStorage.setItem('casual.collab.namePrompted', '1');
      } catch (_) {}
    })();
  `;
}

/**
 * Ensure the thread-comment plugin (+ its facade extension) is mounted on
 * a page. The app idle-loads every lazy group after boot, but a fresh
 * test shouldn't race that — poll until `FRange.addCommentAsync` exists,
 * nudging the idle-load along by reading a comment (a no-op that the
 * shell wires to `ensurePlugin('threadComment')` is not guaranteed, so we
 * just wait for the idle-load to land).
 */
async function waitForCommentFacade(page: Page): Promise<void> {
  await page.waitForFunction(
    () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__univerAPI;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api?.getActiveWorkbook?.()?.getActiveSheet?.();
      return typeof ws?.getRange?.('A1')?.addCommentAsync === 'function';
    },
    null,
    { timeout: 20_000 },
  );
}

async function joinTwoPeerRoom(): Promise<{
  owner: Page;
  joiner: Page;
  cleanup: () => Promise<void>;
}> {
  const ownerCtx = await browser!.newContext();
  const owner = await ownerCtx.newPage();
  await owner.addInitScript({ content: installEnv('Alice') });
  await owner.goto(PROD_BASE);
  await waitForUniver(owner);
  const roomId = await owner.evaluate(async () => {
    const res = await fetch('/api/rooms', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: '{}',
    });
    return ((await res.json()) as { roomId: string }).roomId;
  });
  await owner.goto(`${PROD_BASE}/r/${roomId}`);
  await waitForUniver(owner);
  await expect(owner.getByTestId('presence-avatars')).toBeVisible({ timeout: 10_000 });
  const joinerCtx = await browser!.newContext();
  const joiner = await joinerCtx.newPage();
  await joiner.addInitScript({ content: installEnv('Bob') });
  await joiner.goto(`${PROD_BASE}/r/${roomId}`);
  await waitForUniver(joiner);
  await expect(joiner.getByTestId('presence-avatars')).toBeVisible({ timeout: 10_000 });
  // Owner needs the facade to ADD; joiner gets the plugin lazy-loaded by
  // the bridge on replay, but waiting here makes the post-replay read
  // deterministic.
  await waitForCommentFacade(owner);
  await waitForCommentFacade(joiner);
  return {
    owner,
    joiner,
    cleanup: async () => {
      await ownerCtx.close();
      await joinerCtx.close();
    },
  };
}

/** Read the comment thread text anchored at a given cell on a page. */
async function readCommentAt(page: Page, a1: string): Promise<string[]> {
  return page.evaluate((cell) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook().getActiveSheet();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const comments: any[] = ws.getRange(cell).getComments() ?? [];
    return comments.map((c) => {
      const data = c.getCommentData?.() ?? {};
      return String(data.text?.dataStream ?? '')
        .replace(/[\r\n\t]+/g, ' ')
        .trim();
    });
  }, a1);
}

test('a comment added by one peer appears for the other', async () => {
  const { owner, joiner, cleanup } = await joinTwoPeerRoom();

  // Owner adds a comment anchored at C3.
  const added = await owner.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook().getActiveSheet();
    return ws.getRange('C3').addCommentAsync({ dataStream: 'Sync me across peers\r\n' });
  });
  expect(added, 'owner-side addCommentAsync must succeed').toBeTruthy();

  // Confirm it landed locally first.
  await expect
    .poll(async () => (await readCommentAt(owner, 'C3')).join('|'), { timeout: 5_000 })
    .toContain('Sync me across peers');

  // Then it must propagate to the joiner via the op-log bridge.
  await expect
    .poll(async () => (await readCommentAt(joiner, 'C3')).join('|'), { timeout: 10_000 })
    .toContain('Sync me across peers');

  await cleanup();
});

test('a comment added by the joiner appears for the owner (bidirectional)', async () => {
  const { owner, joiner, cleanup } = await joinTwoPeerRoom();

  const added = await joiner.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook().getActiveSheet();
    return ws.getRange('E5').addCommentAsync({ dataStream: 'From the joiner\r\n' });
  });
  expect(added, 'joiner-side addCommentAsync must succeed').toBeTruthy();

  await expect
    .poll(async () => (await readCommentAt(joiner, 'E5')).join('|'), { timeout: 5_000 })
    .toContain('From the joiner');

  await expect
    .poll(async () => (await readCommentAt(owner, 'E5')).join('|'), { timeout: 10_000 })
    .toContain('From the joiner');

  await cleanup();
});
