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
import { waitForUniver } from './_helpers';

/**
 * Cross-peer drawings propagation.
 *
 * STATUS: GUARDED. Originally skipped — replay threw on the joiner
 * with a json1 apply error from `SheetDrawingService.applyJson1`
 * because the joiner's drawing service had no entry for the unit when
 * the bridge replayed the `sheet.mutation.set-drawing-apply` op. Fix
 * landed in `apps/web/src/collab/bridge.ts` as
 * `ensureUnitDrawingDataReady` — a pre-replay hook that calls
 * `sheetDrawingService.registerDrawingData(unitId, {})` before the
 * first drawing mutation for a unit on each peer. Idempotent —
 * re-registering an existing entry is a no-op.
 *
 * This spec is the regression test that pins the fix.
 */

const PROD_BASE = process.env.PROD_BASE ?? 'http://localhost:3000';

const TINY_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=';

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

async function joinTwoPeerRoom() {
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
  return {
    owner,
    joiner,
    cleanup: async () => {
      await ownerCtx.close();
      await joinerCtx.close();
    },
  };
}

test('image inserted by peer A appears on peer B', async () => {
  const { owner, joiner, cleanup } = await joinTwoPeerRoom();

  await owner.waitForFunction(() => typeof window.__ensurePlugin__ === 'function', null, {
    timeout: 5_000,
  });
  await joiner.waitForFunction(() => typeof window.__ensurePlugin__ === 'function', null, {
    timeout: 5_000,
  });
  await owner.evaluate(async () => {
    await window.__ensurePlugin__!('drawing');
  });
  await joiner.evaluate(async () => {
    await window.__ensurePlugin__!('drawing');
  });

  const drawingId = await owner.evaluate(async (pngDataUrl: string) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = wb.getActiveSheet();
    await ws.insertImage(pngDataUrl, 2, 2, 0, 0);
    const images = ws.getImages();
    if (!images.length) throw new Error('owner insertImage produced no drawing');
    return images[0].getId() as string;
  }, TINY_PNG);

  await joiner.waitForTimeout(3000);

  const peerHasDrawing = await joiner.evaluate((targetId: string) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook();
    const snap = wb.save();
    const drawingResource = (snap.resources ?? []).find(
      (r: { name?: string }) => r?.name === 'SHEET_DRAWING_PLUGIN',
    );
    if (!drawingResource) return false;
    return JSON.stringify(drawingResource).includes(targetId);
  }, drawingId);

  expect(peerHasDrawing).toBe(true);

  await cleanup();
});
