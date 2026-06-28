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
 * Editor-flow co-edit test. Earlier coedit-share/coedit-regression tests
 * used `range.setValue()` programmatically — that bypasses Univer's cell
 * editor entirely and only exercises the SetRangeValuesMutation directly.
 * The user-reported bug is in the EDITOR commit path: type, press Enter,
 * peer briefly sees the value, then the cell reverts to its pre-edit
 * state.
 *
 * This test runs against the docker stack on port 3000 (`pnpm test:e2e`
 * with `PROD_BASE=http://localhost:3000`) and captures what mutations
 * actually fire — see the `__capturedMutations` hook in the page init
 * script.
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

/**
 * Hooks `ICommandService.onMutationExecutedForCollab` on the page so the
 * test can assert what mutations the editor commit actually emits. We
 * read off `window.__capturedMutations` after each action.
 */
function installMutationCapture(displayName: string): string {
  return `
    (function () {
      try {
        localStorage.setItem('casual.collab.displayName', ${JSON.stringify(displayName)});
        localStorage.setItem('casual.collab.namePrompted', '1');
      } catch (_) {}
      window.__capturedMutations = [];
      // Wait until univerAPI mounts, then subscribe via the injector.
      var hookUp = function () {
        var api = window.__univerAPI;
        if (!api) return false;
        try {
          var injector = api._injector;
          if (!injector) return false;
          // Univer's ICommandService identifier — string-keyed lookup
          // doesn't work; pull through the facade's running instance.
          var cs = null;
          for (var k in api) { /* hack-fallback if injector path changes */ }
          // The bridge does the same import; we mirror it via the global
          // we already store at mount time.
          if (!window.__ICommandService) return false;
          cs = injector.get(window.__ICommandService);
          if (!cs || typeof cs.onMutationExecutedForCollab !== 'function') return false;
          cs.onMutationExecutedForCollab(function (info, options) {
            try {
              window.__capturedMutations.push({
                id: info.id,
                t: Date.now(),
                fromCollab: !!(options && options.fromCollab),
                paramsSummary: summarise(info.params),
              });
            } catch (e) { /* swallow */ }
          });
          return true;
        } catch (e) { return false; }
      };
      var attempt = function () { if (!hookUp()) setTimeout(attempt, 100); };
      attempt();

      function summarise(p) {
        if (!p || typeof p !== 'object') return null;
        var out = { keys: Object.keys(p).slice(0, 8) };
        if (typeof p.unitId === 'string') out.unitId = p.unitId;
        if (typeof p.subUnitId === 'string') out.subUnitId = p.subUnitId;
        if (p.cellValue && typeof p.cellValue === 'object') {
          out.cells = [];
          for (var r in p.cellValue) {
            var row = p.cellValue[r];
            if (row && typeof row === 'object') {
              for (var c in row) {
                out.cells.push({ r: +r, c: +c, v: row[c] && row[c].v });
              }
            }
          }
        }
        return out;
      }
    })();
  `;
}

test('editor commit propagates the typed value AND it does not revert', async () => {
  // ─ owner ─
  const ownerCtx = await browser!.newContext();
  const owner = await ownerCtx.newPage();
  await owner.addInitScript({ content: installMutationCapture('Alice') });
  await owner.goto(PROD_BASE);
  await waitForUniver(owner);

  // Make ICommandService visible to the init-script's hook.
  await owner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = window.__univerAPI as any;
    // The bridge already imports ICommandService; we re-export it on
    // window so the page init-script can grab the identity (Univer DI
    // uses identifier-by-reference, so we have to use Univer's own).
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__ICommandService = api._injector?.get?.bind(api._injector);
    void api;
  });

  // Create open room via API, navigate as owner.
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

  // ─ joiner ─
  const joinerCtx = await browser!.newContext();
  const joiner = await joinerCtx.newPage();
  joiner.on('console', (m) => {
    const t = m.text();
    if (t.includes('[collab]') || t.includes('[bridge]') || t.includes('[replay]'))
      console.log('JOINER CONSOLE:', t);
  });
  await joiner.addInitScript({ content: installMutationCapture('Bob') });
  await joiner.goto(`${PROD_BASE}/r/${roomId}`);
  await waitForUniver(joiner);
  await expect(joiner.getByTestId('presence-avatars')).toBeVisible({ timeout: 10_000 });
  // Both peers see two avatars.
  await expect(joiner.getByTestId('presence-avatar')).toHaveCount(2, { timeout: 10_000 });
  await expect(owner.getByTestId('presence-avatar')).toHaveCount(2, { timeout: 10_000 });

  // ─ owner types via the EDITOR (not setValue) ─
  // Focus B5 by activating its range first (so the canvas knows where to
  // open the editor), then double-click to enter edit mode, type, Enter.
  await owner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = window.__univerAPI as any;
    api.getActiveWorkbook().getActiveSheet().getRange('B5').activate();
  });
  // Univer's main canvas is a single element; double-click positions
  // need pixel coords, but we can use the F2 keystroke shortcut to enter
  // edit mode on the active cell (same as Excel).
  const canvas = owner.locator('canvas[id^="univer-sheet-main-canvas_"]').first();
  await canvas.focus();
  await owner.keyboard.press('F2');
  // Brief settle so the editor mounts.
  await owner.waitForTimeout(150);
  await owner.keyboard.type('hello-from-alice', { delay: 25 });
  await owner.keyboard.press('Enter');

  // Diagnostic: wait a bit then dump both sides' state regardless of
  // whether propagation worked.
  await joiner.waitForTimeout(3000);
  const diag = await joiner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook().getActiveSheet();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const provider = (window as any).__hocuspocusProvider;
    return {
      b5: ws.getRange('B5').getValue(),
      b5raw: ws.getRange('B5').getCellData(),
      unitId: api.getActiveWorkbook().getId(),
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      logLen: provider?.document?.getArray?.('ops')?.length,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      logEntries: provider?.document?.getArray?.('ops')?.toArray?.(),
      providerStatus: provider?.status,
      synced: provider?.synced,
    };
  });
  const ownerDiag = await owner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook();
    const ws = wb.getActiveSheet();
    return {
      b5: ws.getRange('B5').getValue(),
      b5raw: ws.getRange('B5').getCellData(),
      unitId: wb.getId(),
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      sheetIds: wb.getSheets().map((s: any) => s.getSheetId()),
    };
  });
  const joinerSheets = await joiner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return wb.getSheets().map((s: any) => s.getSheetId());
  });
  console.log('JOINER sheetIds:', joinerSheets);
  console.log('OWNER  B5:', JSON.stringify(ownerDiag));
  console.log('JOINER B5:', JSON.stringify(diag));

  expect(diag.b5, 'joiner should see the typed value').toBe('hello-from-alice');

  // Wait long enough for any post-commit revert mutation to land.
  // The reported symptom is "changes first to remote participant but on
  // changing cursor it reverts" — give it 2 s, which is well past any
  // microtask/queueMicrotask flush + Yjs round-trip.
  await joiner.waitForTimeout(2000);

  const stickValue = await joiner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook().getActiveSheet();
    return ws.getRange('B5').getValue();
  });
  expect(stickValue, 'value should still be the typed text 2 s after commit').toBe(
    'hello-from-alice',
  );

  // For diagnostic / regression — dump every mutation captured on both
  // sides so a future failure has a trail to follow.
  const ownerMuts = await owner.evaluate(() => (window as unknown as { __capturedMutations: unknown[] }).__capturedMutations);
  const joinerMuts = await joiner.evaluate(() => (window as unknown as { __capturedMutations: unknown[] }).__capturedMutations);
  console.log('[owner mutations]', JSON.stringify(ownerMuts, null, 2));
  console.log('[joiner mutations]', JSON.stringify(joinerMuts, null, 2));

  await ownerCtx.close();
  await joinerCtx.close();
});
