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
 * Verifies cross-peer conditional formatting:
 *   - peer A adds a "highlight if > 5" rule on A1:A5
 *   - peer B sees the rule in their ConditionalFormattingRuleModel
 *
 * Doesn't assert the actual cell rendering (canvas pixels are
 * brittle across runners); the rule-model presence proves the
 * `sheet.mutation.add-conditional-rule` mutation propagated.
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

test('conditional formatting rule added by peer A appears on peer B', async () => {
  const { owner, joiner, cleanup } = await joinTwoPeerRoom();

  // Owner adds a CF rule via sheet.command.add-conditional-rule. The
  // command lazy-loads the plugin, so wait for it to mount first.
  const cfId = await owner.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook();
    const id = 'cf-' + Math.random().toString(36).slice(2, 9);
    await api.executeCommand('sheet.command.add-conditional-rule', {
      unitId: wb.getId(),
      subUnitId: wb.getActiveSheet().getSheetId(),
      rule: {
        cfId: id,
        ranges: [{ startRow: 0, startColumn: 0, endRow: 4, endColumn: 0 }],
        rule: {
          type: 'highlightCell',
          subType: 'number',
          operator: 'greaterThan',
          value: 5,
          style: { bg: { rgb: '#fff59d' } },
        },
        stopIfTrue: false,
      },
    });
    return id;
  });
  // CF plugin lazy-loads; allow extra time on first invocation.
  await joiner.waitForTimeout(2500);

  // Verify via wb.save() — Univer serialises CF rules into the
  // resources channel of every save. If the mutation handler ran on
  // the joiner, the rule shows up there. Avoids the brittleness of
  // poking the injector for the rule-model instance.
  const peerHasRule = await joiner.evaluate((targetCfId: string) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook();
    const snap = wb.save();
    const cfResource = (snap.resources ?? []).find(
      (r: { name?: string; data?: string }) => r?.name === 'SHEET_CONDITIONAL_FORMATTING_PLUGIN',
    );
    if (!cfResource) {
      return {
        ok: false,
        reason: `no SHEET_CONDITIONAL_FORMATTING_PLUGIN resource in ${(snap.resources ?? []).length} resources: ${(snap.resources ?? []).map((r: { name?: string }) => r?.name).join(',')}`,
      };
    }
    // Resource data is usually a JSON string keyed by subUnitId → rules[].
    let parsed: unknown;
    try {
      parsed = typeof cfResource.data === 'string' ? JSON.parse(cfResource.data) : cfResource.data;
    } catch {
      return { ok: false, reason: 'cf resource data not parseable' };
    }
    const blob = JSON.stringify(parsed ?? '');
    return {
      ok: blob.includes(targetCfId),
      reason: blob.includes(targetCfId) ? 'present' : `not in ${blob.slice(0, 200)}`,
    };
  }, cfId);

  expect(peerHasRule.ok, `peer should have CF rule ${cfId}: ${peerHasRule.reason}`).toBe(true);

  await cleanup();
});
