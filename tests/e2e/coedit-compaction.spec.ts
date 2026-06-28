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
 * Pipeline Stage 6 — op-log compaction. Drives the bridge past the
 * compaction threshold, then confirms:
 *
 *   1. The Y.Array log shrinks atomically (200+ ops → 1 snapshot
 *      record + ~0 trailing ops).
 *   2. A second browser joining AFTER compaction reconstructs the
 *      workbook from the snapshot record (no missing cells).
 *   3. Subsequent edits keep syncing as usual.
 *
 * Runs against its own Hocuspocus on a dedicated port — mirrors the
 * pattern in coedit.spec / coedit-share.spec so the suite can fan
 * out without port collisions.
 */

test.describe.configure({ mode: 'serial', retries: 3 });
// This spec drives a heavy collab flow (220 mutations → op-log compaction →
// a second browser context replaying the snapshot) against the Vite dev server
// under 2-worker CI contention. The per-step waits below are deliberately
// generous: the flow is correct (passes locally in ~6s and on faster CI runs)
// but on a slow/contended runner any single step can drift past a tight budget,
// which is what made this the lone flaky failure in main's e2e job. Give the
// whole test plenty of room rather than tuning one wait at a time.
test.setTimeout(360_000);

const SERVER_PORT = 3060;
const WS_URL = `ws://127.0.0.1:${SERVER_PORT}/yjs`;

let serverProc: ChildProcess | null = null;
let browser: Browser | null = null;

test.beforeAll(async () => {
  serverProc = spawn('pnpm', ['--filter', '@casualoffice/collab', 'exec', 'tsx', 'src/index.ts'], {
    env: { ...process.env, PORT: String(SERVER_PORT), HOST: '127.0.0.1' },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  await new Promise<void>((resolveReady, reject) => {
    const timer = setTimeout(() => reject(new Error('server boot timed out')), 60_000);
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

/** Reach into the bridge module from the page context and trip the
 *  compaction immediately, regardless of the 30s timer. The bridge
 *  module isn't exported globally; we go through Yjs's awareness
 *  state to find the designated writer, then dispatch enough work
 *  to cross the 200-op threshold, then wait for the log shrink. */
async function runCompactionOnce(
  page: import('@playwright/test').Page,
  roomId: string,
): Promise<{ before: number; after: number }> {
  // 220 setRangeValues mutations from one browser — past the 200-op
  // compaction threshold.
  await page.evaluate(async () => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    for (let i = 0; i < 220; i++) {
      ws.getRange(i, 0).setValue({ v: `row-${i}` });
    }
  });
  // Wait for the bridge's per-microtask batch flush to land, then
  // poll until the log is observably non-empty.
  await page.waitForFunction(
    () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const probe = (window as any).__bridgeLogLength;
      return typeof probe === 'function' && probe() >= 200;
    },
    null,
    { timeout: 90_000 },
  );
  const before = await page.evaluate(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    () => ((window as any).__bridgeLogLength as () => number)(),
  );
  // Force compaction now instead of waiting for the 30s interval.
  await page.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const force = (window as any).__bridgeForceCompact as (() => void) | undefined;
    if (typeof force === 'function') force();
  });
  // Wait for the log to collapse.
  await page.waitForFunction(
    () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const probe = (window as any).__bridgeLogLength;
      return typeof probe === 'function' && probe() <= 5;
    },
    null,
    { timeout: 90_000 },
  );
  const after = await page.evaluate(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    () => ((window as any).__bridgeLogLength as () => number)(),
  );
  void roomId;
  return { before, after };
}

test('op log compacts after threshold and joiners replay the snapshot', async ({ baseURL }) => {
  const roomId = `t-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
  const url = `${baseURL}/r/${roomId}`;

  const a = await browser!.newContext();
  const pageA = await a.newPage();
  await pageA.addInitScript((wsUrl: string) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__COLLAB_WS_URL__ = wsUrl;
    try {
      localStorage.setItem('casual.collab.namePrompted', '1');
    } catch {
      /* private mode */
    }
  }, WS_URL);

  await pageA.goto(url);
  await waitForUniver(pageA);

  const { before, after } = await runCompactionOnce(pageA, roomId);
  expect(before).toBeGreaterThanOrEqual(200);
  expect(after).toBeLessThanOrEqual(5);

  // Joiner connects AFTER compaction. Should see all 220 rows from
  // the snapshot record without ever touching the original ops.
  const b = await browser!.newContext();
  const pageB = await b.newPage();
  await pageB.addInitScript((wsUrl: string) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__COLLAB_WS_URL__ = wsUrl;
    try {
      localStorage.setItem('casual.collab.namePrompted', '1');
    } catch {
      /* private mode */
    }
  }, WS_URL);
  await pageB.goto(url);
  await waitForUniver(pageB);

  // Poll for the data to land — workbook replace runs through the
  // unit-swap effect, so there's a small async gap after sync.
  await pageB.waitForFunction(
    () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()?.getActiveSheet();
      return (
        ws?.getRange(0, 0).getValue() === 'row-0' && ws?.getRange(219, 0).getValue() === 'row-219'
      );
    },
    null,
    // The joiner loads the full app cold (second browser context) before the
    // snapshot replay lands. On a slow/cold CI runner against the Vite dev
    // server, 10s was too tight and flaked here; 30s is comfortably within the
    // 120s test budget and reflects real worst-case CI boot+replay time.
    { timeout: 90_000 },
  );

  await a.close();
  await b.close();
});
