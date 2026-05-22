import { expect, test, chromium, type Browser } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Cross-peer data-validation propagation. Same pattern as the CF
 * regression: peer A adds a list-validation rule on A1:A5; peer B's
 * workbook should hold the rule in its serialised resources.
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

test('data-validation rule added by peer A appears on peer B', async () => {
  const { owner, joiner, cleanup } = await joinTwoPeerRoom();

  const ruleId = await owner.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook();
    const uid = 'dv-' + Math.random().toString(36).slice(2, 9);
    await api.executeCommand('sheet.command.addDataValidation', {
      unitId: wb.getId(),
      subUnitId: wb.getActiveSheet().getSheetId(),
      rule: {
        uid,
        type: 'list',
        formula1: 'apple,banana,cherry',
        ranges: [{ startRow: 0, startColumn: 0, endRow: 4, endColumn: 0 }],
      },
    });
    return uid;
  });
  await joiner.waitForTimeout(2500);

  const peerHasRule = await joiner.evaluate((targetUid: string) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook();
    const snap = wb.save();
    const dvResource = (snap.resources ?? []).find(
      (r: { name?: string }) => r?.name === 'SHEET_DATA_VALIDATION_PLUGIN',
    );
    if (!dvResource) {
      return {
        ok: false,
        reason: `no SHEET_DATA_VALIDATION_PLUGIN in resources: ${(snap.resources ?? []).map((r: { name?: string }) => r?.name).join(',')}`,
      };
    }
    const blob = JSON.stringify(dvResource);
    return {
      ok: blob.includes(targetUid),
      reason: blob.includes(targetUid) ? 'present' : `not in ${blob.slice(0, 200)}`,
    };
  }, ruleId);

  expect(peerHasRule.ok, `peer should have DV rule ${ruleId}: ${peerHasRule.reason}`).toBe(true);

  await cleanup();
});
