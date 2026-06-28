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
 * Regressions for the cursor coordinate math fixes:
 *   - #11/19d frozen-pane awareness: peer cursors in frozen rows must
 *     stay at a constant viewport-Y even as the joiner scrolls past.
 *   - #11 zoom-aware: at non-100% zoom the cursor must scale with the
 *     canvas, not drift proportionally to (1 - zoom).
 *
 * Pixel assertions are softened to tolerance windows because canvas
 * rendering varies subpixel-by-subpixel across CI runners; the goal
 * is to catch ORDER-OF-MAGNITUDE drift, not perfect alignment.
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

async function joinTwoPeerRoom(): Promise<{
  owner: import('@playwright/test').Page;
  joiner: import('@playwright/test').Page;
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

  return {
    owner,
    joiner,
    cleanup: async () => {
      await ownerCtx.close();
      await joinerCtx.close();
    },
  };
}

test('peer cursor scales with zoom on the receiving end', async () => {
  const { owner, joiner, cleanup } = await joinTwoPeerRoom();

  // Owner parks the selection on a known cell — F10 is far enough
  // from the header gutter that the coordinate math has room to be
  // visibly wrong if it forgets the gutter or the zoom.
  await owner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    api.getActiveWorkbook().getActiveSheet().getRange('F10').activate();
  });

  // Wait for the peer cursor to appear on the joiner side.
  await expect(joiner.getByTestId('presence-cursor')).toBeVisible({ timeout: 5_000 });

  // Capture cursor rect at 100% zoom.
  const before = await joiner.evaluate(() => {
    const el = document.querySelector('[data-testid="presence-cursor"]');
    if (!el) return null;
    const r = (el as HTMLElement).getBoundingClientRect();
    return { left: r.left, top: r.top, width: r.width, height: r.height };
  });
  expect(before).not.toBeNull();

  // Zoom the joiner to 50%.
  await joiner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook();
    api.executeCommand('sheet.command.set-zoom-ratio', {
      zoomRatio: 0.5,
      unitId: wb.getId(),
      subUnitId: wb.getActiveSheet().getSheetId(),
    });
  });
  await joiner.waitForTimeout(500);

  const after = await joiner.evaluate(() => {
    const el = document.querySelector('[data-testid="presence-cursor"]');
    if (!el) return null;
    const r = (el as HTMLElement).getBoundingClientRect();
    return { left: r.left, top: r.top, width: r.width, height: r.height };
  });
  expect(after).not.toBeNull();

  // The smoking-gun signal that the zoom math is wired: the cursor
  // dimensions must shrink with zoom — pre-fix the width/height
  // ignored zoom entirely and stayed at full size. (We don't pin the
  // exact ratio because the joiner's viewport scroll also reflows on
  // zoom change, but the cursor RECT itself should be smaller.)
  expect(after!.width).toBeLessThan(before!.width);
  expect(after!.height).toBeLessThan(before!.height);
  // And the size delta should be substantial — anything <30% shrink
  // would suggest the zoom factor isn't getting applied.
  expect(after!.width / before!.width).toBeLessThan(0.7);

  await cleanup();
});

test('peer cursor in a frozen row stays put when the joiner scrolls', async () => {
  const { owner, joiner, cleanup } = await joinTwoPeerRoom();

  // Both peers share a single workbook; setting freeze on the owner
  // propagates via the SetFrozenMutation (which IS in
  // SYNCED_MUTATIONS — see bridge.ts). Freeze rows 0-1 + cols 0-1 so
  // a cell in the frozen quadrant has a sensible test target.
  await owner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook();
    api.executeCommand('sheet.command.set-frozen', {
      unitId: wb.getId(),
      subUnitId: wb.getActiveSheet().getSheetId(),
      startRow: 2,
      startColumn: 2,
      ySplit: 2,
      xSplit: 2,
    });
  });
  await joiner.waitForTimeout(500);

  // Owner parks selection on a frozen-row cell (row 0). We use B1 —
  // inside both the frozen ySplit (row < 2) and the frozen xSplit
  // (col < 2), so it sits in the top-left frozen quadrant that
  // doesn't scroll on either axis.
  await owner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    api.getActiveWorkbook().getActiveSheet().getRange('B1').activate();
  });
  await expect(joiner.getByTestId('presence-cursor')).toBeVisible({ timeout: 5_000 });

  const cursorTopBefore = await joiner.evaluate(() => {
    const el = document.querySelector('[data-testid="presence-cursor"]');
    return el ? (el as HTMLElement).getBoundingClientRect().top : null;
  });
  expect(typeof cursorTopBefore).toBe('number');

  // Scroll the joiner down by a big chunk. Use Univer's scrollToCell
  // to jump to row 200 — guarantees the non-frozen viewport has moved.
  await joiner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    api.executeCommand('sheet.operation.scroll-to-cell', { range: { startRow: 200, startColumn: 5, endRow: 200, endColumn: 5 } });
  });
  await joiner.waitForTimeout(500);

  const cursorTopAfter = await joiner.evaluate(() => {
    const el = document.querySelector('[data-testid="presence-cursor"]');
    return el ? (el as HTMLElement).getBoundingClientRect().top : null;
  });
  expect(typeof cursorTopAfter).toBe('number');

  // The frozen-row cursor should stay within a small tolerance of its
  // pre-scroll position. Pre-fix behaviour: cursor would shift up by
  // roughly (rows scrolled × default row height) → hundreds of pixels.
  // Tolerance of 8 px allows for layout settling but reliably catches
  // a "the cursor drifted with the scrollable area" regression.
  const drift = Math.abs(cursorTopAfter! - cursorTopBefore!);
  expect(drift).toBeLessThan(8);

  await cleanup();
});
