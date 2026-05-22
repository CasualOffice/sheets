import { expect, test, chromium, type Browser } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Regression for `WorkbookEditablePermission` being flipped on
 * `role=view` joiners. Earlier code only filtered outbound mutations
 * client-side — the editor still opened locally, the user typed,
 * nothing synced, and they discovered the disconnect only when
 * comparing notes with the owner. Now Univer itself refuses the edit.
 *
 * Runs against the docker prod stack via playwright.docker.config.ts.
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

test('view-role joiner cannot mutate cells via the editor', async () => {
  // Owner creates an open room.
  const ownerCtx = await browser!.newContext();
  const owner = await ownerCtx.newPage();
  await owner.addInitScript({ content: installEnv('Owner') });
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

  // View-role joiner.
  const viewerCtx = await browser!.newContext();
  const viewer = await viewerCtx.newPage();
  await viewer.addInitScript({ content: installEnv('Reader') });
  await viewer.goto(`${PROD_BASE}/r/${roomId}?role=view`);
  await waitForUniver(viewer);
  await expect(viewer.getByTestId('view-only-banner')).toBeVisible({ timeout: 10_000 });

  // Attempt to type via the editor. F2 should be refused by Univer
  // because WorkbookEditablePermission is false. Even if Univer opens
  // a partial editor, the commit mutation should not actually update
  // the cell because the permission check blocks set-range-values.
  await viewer.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    api.getActiveWorkbook().getActiveSheet().getRange('C7').activate();
  });
  const canvas = viewer.locator('canvas[id^="univer-sheet-main-canvas_"]').first();
  await canvas.focus();
  await viewer.keyboard.press('F2');
  await viewer.waitForTimeout(150);
  await viewer.keyboard.type('should-not-stick', { delay: 15 });
  await viewer.keyboard.press('Enter');
  await viewer.waitForTimeout(500);

  // Cell stays empty locally.
  const viewerC7 = await viewer.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (api.getActiveWorkbook().getActiveSheet() as any).getRange('C7').getValue();
  });
  expect(viewerC7, 'view-only user must not write to their own cell').toBeFalsy();

  // And the owner doesn't see it either — extra belt-and-braces.
  await owner.waitForTimeout(500);
  const ownerC7 = await owner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (api.getActiveWorkbook().getActiveSheet() as any).getRange('C7').getValue();
  });
  expect(ownerC7, 'view-only edits must never propagate to peers').toBeFalsy();

  await ownerCtx.close();
  await viewerCtx.close();
});
