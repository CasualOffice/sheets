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
import { mainCanvas, waitForUniver } from './_helpers';

// Single-worker — shares one Hocuspocus instance via beforeAll, same
// pattern as coedit.spec.ts. Running parallel workers would collide on
// the fixed port and the second one would exit before tests start.
test.describe.configure({ mode: 'serial' });

/**
 * Charts P5 — collab sync. Two browser contexts join the same room.
 * Inserting a chart in A appears in B's overlay + Charts panel within
 * a few seconds. Renaming and deleting also propagate.
 */

const SERVER_PORT = 3061;
const WS_URL = `ws://127.0.0.1:${SERVER_PORT}/yjs`;

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

test('chart insert in browser A appears in browser B within 3s', async ({ baseURL }) => {
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

  // Seed shared data via A — it will sync to B through the existing
  // bridge before we insert the chart.
  await pageA.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Region' });
    ws.getRange('B1').setValue({ v: 'Q1' });
    ws.getRange('C1').setValue({ v: 'Q2' });
    ws.getRange('A2').setValue({ v: 'N' });
    ws.getRange('B2').setValue({ v: 100 });
    ws.getRange('C2').setValue({ v: 120 });
    ws.getRange('A3').setValue({ v: 'S' });
    ws.getRange('B3').setValue({ v: 80 });
    ws.getRange('C3').setValue({ v: 95 });
    ws.getRange('A1:C3').activate();
  });
  await pageB.waitForFunction(
    () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange('B2').getValue() === 100;
    },
    { timeout: 5_000 },
  );

  // Insert a chart on A via the menu, with a real range selected.
  await mainCanvas(pageA).first().click({ position: { x: 100, y: 100 } });
  await pageA.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1:C3').activate();
  });
  await pageA.getByTestId('menubar-insert').click();
  await pageA.getByTestId('menu-item-insert-chart').click();
  await pageA.getByTestId('insert-chart-confirm').click();
  await expect(pageA.getByTestId('chart-overlay')).toBeVisible({ timeout: 5_000 });

  // B should grow its chart overlay within a few seconds.
  await expect(pageB.getByTestId('chart-overlay')).toBeVisible({ timeout: 5_000 });

  // Open B's Charts panel — should also list "Chart 1".
  await pageB.getByTestId('menubar-view').click();
  await pageB.getByTestId('menu-item-charts-panel').click();
  await expect(pageB.getByTestId('charts-panel').getByText('Chart 1')).toBeVisible();

  // Rename on A via the right-click menu → propagates to B's panel.
  await pageA.getByTestId('chart-overlay').click({ button: 'right' });
  await pageA.getByTestId('chart-context-rename').click();
  await pageA.getByTestId('chart-context-rename-input').fill('Synced revenue');
  await pageA.getByTestId('chart-context-rename-input').press('Enter');
  await expect(pageB.getByTestId('charts-panel').getByText('Synced revenue')).toBeVisible({
    timeout: 5_000,
  });

  // Delete on B via the panel → removes A's overlay.
  await pageB.getByLabel(/^Delete Synced revenue$/).click();
  await expect(pageA.getByTestId('chart-overlay')).toHaveCount(0, { timeout: 5_000 });

  await a.close();
  await b.close();
});
