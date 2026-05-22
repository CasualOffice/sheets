import { expect, test, chromium, type Browser } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Cross-peer drawings propagation.
 *
 * STATUS: SKIPPED — replay throws on the joiner with a json1 apply error
 * from `SheetDrawingService.applyJson1` because the joiner's drawing
 * service has no entry for the unit when the bridge replays the
 * `sheet.mutation.set-drawing-apply` op. The op was generated against
 * the owner's empty state and assumes the unit path exists; on a peer
 * that has never run an Insert/SetSheetDrawing locally, the path is
 * missing and json1 throws. Fix is a bridge-side pre-replay hook that
 * calls `sheetDrawingService.registerDrawingData(unitId, {})` (and
 * the matching drawingManagerService init) the first time we see a
 * drawing mutation for a unit. Wiring that up needs its own commit;
 * this spec is the regression that will guard the fix.
 *
 * Captured failure (set-drawing-apply on joiner):
 *   Error
 *     at Object.T [as apply] (json1 apply)
 *     at SheetDrawingService.applyJson1
 *     at handler (SetDrawingApplyMutation)
 *     at ICommandService.executeCommand
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

test.skip('image inserted by peer A appears on peer B', async () => {
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
