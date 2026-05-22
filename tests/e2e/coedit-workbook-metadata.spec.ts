import { expect, test, chromium, type Browser } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Cross-peer propagation for the workbook / worksheet metadata mutations
 * that were added to the bridge's SYNCED_MUTATIONS allowlist. Each one is
 * rarely changed mid-session but, before being allow-listed, silently
 * stayed local — confusing in a shared room. The test fires the command
 * on peer A and reads peer B's serialised snapshot to assert the state
 * actually crossed the wire.
 *
 *   - sheet.command.set-workbook-name        → workbook.name
 *   - sheet.command.toggle-gridlines         → worksheet.showGridlines
 *   - sheet.command.set-worksheet-row-count  → worksheet.rowCount
 *
 * (RTL flip is not tested — the underlying command/mutation pair is
 *  exported but never registered by any plugin in @univerjs/sheets@0.22.1.)
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

test('workbook rename propagates to peers', async () => {
  const { owner, joiner, cleanup } = await joinTwoPeerRoom();
  const targetName = 'Renamed-' + Math.random().toString(36).slice(2, 8);
  await owner.evaluate(async (newName: string) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook();
    await api.executeCommand('sheet.command.set-workbook-name', {
      unitId: wb.getId(),
      name: newName,
    });
  }, targetName);

  // Local apply sanity check.
  const ownerName = await owner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (api.getActiveWorkbook().save() as any).name;
  });
  expect(ownerName, 'rename must apply locally first').toBe(targetName);

  await joiner.waitForTimeout(1500);
  const peerName = await joiner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (api.getActiveWorkbook().save() as any).name;
  });
  expect(peerName).toBe(targetName);
  await cleanup();
});

test('toggle-gridlines propagates to peers', async () => {
  const { owner, joiner, cleanup } = await joinTwoPeerRoom();
  // Capture default first, then toggle once, then assert peer flipped to
  // the same new state. We don't pin a specific value because the default
  // could change between Univer versions; we only assert "both peers agree
  // AND the value actually changed."
  const before = await owner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook();
    const sid = wb.getActiveSheet().getSheetId();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return { sid, show: (wb.save() as any).sheets[sid].showGridlines };
  });

  await owner.evaluate(
    async ({ sid, target }: { sid: string; target: number }) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__univerAPI;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb: any = api.getActiveWorkbook();
      // The command's `params.showGridlines` is an idempotency guard:
      // if it equals the CURRENT value the command early-returns false
      // (no toggle). To actually flip we must pass the value we want
      // AFTER the toggle — i.e. the inverse of current. The handler
      // computes the new value itself and pushes the mutation.
      await api.executeCommand('sheet.command.toggle-gridlines', {
        unitId: wb.getId(),
        subUnitId: sid,
        showGridlines: target,
      });
    },
    { sid: before.sid, target: before.show === 1 ? 0 : 1 },
  );

  const ownerAfter = await owner.evaluate((sid: string) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (api.getActiveWorkbook().save() as any).sheets[sid].showGridlines;
  }, before.sid);
  expect(ownerAfter, 'gridlines must flip locally first').not.toBe(before.show);

  await joiner.waitForTimeout(1500);
  const peerAfter = await joiner.evaluate((sid: string) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (api.getActiveWorkbook().save() as any).sheets[sid].showGridlines;
  }, before.sid);
  expect(peerAfter).toBe(ownerAfter);
  await cleanup();
});

test('row-count grow propagates to peers', async () => {
  const { owner, joiner, cleanup } = await joinTwoPeerRoom();
  const sid = await owner.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook();
    const s = wb.getActiveSheet().getSheetId();
    await api.executeCommand('sheet.command.set-worksheet-row-count', {
      unitId: wb.getId(),
      subUnitId: s,
      rowCount: 5000,
    });
    return s;
  });

  const ownerRows = await owner.evaluate((id: string) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (api.getActiveWorkbook().save() as any).sheets[id].rowCount;
  }, sid);
  expect(ownerRows).toBe(5000);

  await joiner.waitForTimeout(1500);
  const peerRows = await joiner.evaluate((id: string) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (api.getActiveWorkbook().save() as any).sheets[id].rowCount;
  }, sid);
  expect(peerRows).toBe(5000);
  await cleanup();
});
