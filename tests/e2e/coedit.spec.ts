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

import { expect, test, chromium, type Browser, type BrowserContext } from '@playwright/test';
import { spawn, type ChildProcess } from 'node:child_process';
import { waitForUniver } from './_helpers';

// Both tests share the same Hocuspocus instance (booted in beforeAll on a
// fixed port). Running in parallel workers would try to bind the same port
// twice → second worker's server exits 1 before any test runs. Serial mode
// keeps the file on a single worker.
test.describe.configure({ mode: 'serial' });

/**
 * Co-editing spike — Spike A from PLAN.md, hardened slightly. Two browser
 * contexts open the same room URL. Edit A1 in context 1; within 1 s, A1
 * shows the new value in context 2.
 *
 * Setup: we spin up the Hocuspocus + Fastify server ourselves on a fresh
 * port (different from the playwright dev server's :5273) so the test is
 * self-contained — no docker, no shared compose state, no flakes from
 * port collisions with `pnpm dev:server`.
 */

const SERVER_PORT = 3055;
const WS_URL = `ws://127.0.0.1:${SERVER_PORT}/yjs`;

let serverProc: ChildProcess | null = null;
let browser: Browser | null = null;

test.beforeAll(async () => {
  // Run the server via pnpm filter so tsx is resolved from
  // apps/server/node_modules — direct `pnpm exec tsx` from the repo root
  // fails because the workspace root doesn't have tsx installed.
  serverProc = spawn(
    'pnpm',
    ['--filter', '@casualoffice/collab', 'exec', 'tsx', 'src/index.ts'],
    {
      env: { ...process.env, PORT: String(SERVER_PORT), HOST: '127.0.0.1' },
      stdio: ['ignore', 'pipe', 'pipe'],
    },
  );

  await new Promise<void>((resolveReady, reject) => {
    const timer = setTimeout(
      () => reject(new Error('server boot timed out')),
      15_000,
    );
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

test('two browsers in the same room sync a single cell edit', async ({ baseURL }) => {
  // Each context = its own browser session: separate origin storage,
  // separate Univer instance, separate Yjs Doc. Same Hocuspocus room.
  const roomId = `t-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
  const url = `${baseURL}/r/${roomId}`;

  const a: BrowserContext = await browser!.newContext();
  const b: BrowserContext = await browser!.newContext();
  const pageA = await a.newPage();
  const pageB = await b.newPage();

  // Point both pages at our standalone server's WS port. Set BEFORE
  // navigation so CollabDriver picks up the override at mount time.
  for (const p of [pageA, pageB]) {
    await p.addInitScript((url) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (window as any).__COLLAB_WS_URL__ = url;
      // Skip the first-time name prompt — this test doesn't exercise it.
      try {
        localStorage.setItem('casual.collab.namePrompted', '1');
      } catch {
        /* private mode */
      }
    }, WS_URL);
  }

  await pageA.goto(url);
  await waitForUniver(pageA);
  await pageB.goto(url);
  await waitForUniver(pageB);

  // Type into A1 on page A.
  const HELLO = `hello-${Date.now()}`;
  await pageA.evaluate((v) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v });
  }, HELLO);

  // Within 3 s, page B should show the same value in A1.
  await pageB.waitForFunction(
    (expected) => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange('A1').getValue() === expected;
    },
    HELLO,
    { timeout: 3_000 },
  );

  // Send an edit the OTHER way too — confirms the bridge is symmetric.
  const BACK = `back-${Date.now()}`;
  await pageB.evaluate((v) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('B2').setValue({ v });
  }, BACK);

  await pageA.waitForFunction(
    (expected) => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange('B2').getValue() === expected;
    },
    BACK,
    { timeout: 3_000 },
  );

  await a.close();
  await b.close();
});

test('style + structural mutations propagate across peers', async ({ baseURL }) => {
  const roomId = `t-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
  const url = `${baseURL}/r/${roomId}`;

  const a = await browser!.newContext();
  const b = await browser!.newContext();
  const pageA = await a.newPage();
  const pageB = await b.newPage();
  for (const p of [pageA, pageB]) {
    await p.addInitScript((wsUrl) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (window as any).__COLLAB_WS_URL__ = wsUrl;
      try {
        localStorage.setItem('casual.collab.namePrompted', '1');
      } catch {
        /* private mode */
      }
    }, WS_URL);
  }

  await pageA.goto(url);
  await waitForUniver(pageA);
  await pageB.goto(url);
  await waitForUniver(pageB);

  // 1) Bold via the toolbar on A → B sees the cell as bold.
  await pageA.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('C3').setValue({ v: 'styled' });
    ws.getRange('C3').activate();
  });
  await pageA.getByTestId('ribbon-btn-bold').click();

  await pageB.waitForFunction(
    () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb: any = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = wb.getActiveSheet();
      const cd = ws.getRange('C3').getCellData();
      const styleRef = cd?.s;
      const style =
        typeof styleRef === 'string'
          ? wb.getWorkbook().getStyles().get(styleRef)
          : (styleRef ?? null);
      return style?.bl === 1;
    },
    null,
    { timeout: 5_000 },
  );

  // 2) Insert a new sheet on B → A sees the same sheet count.
  await pageB.evaluate(() => {
    const api = window.__univerAPI!;
    api.getActiveWorkbook()!.insertSheet();
  });
  await pageA.waitForFunction(
    () => {
      const api = window.__univerAPI!;
      return api.getActiveWorkbook()!.getSheets().length >= 2;
    },
    null,
    { timeout: 5_000 },
  );

  await a.close();
  await b.close();
});
