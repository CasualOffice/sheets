import { expect, test } from '@playwright/test';

/**
 * Off-main formula compute: `<CasualSheets formula={{ worker }}>` registers the
 * formula plugins with `notExecuteFormula` + `UniverRPCMainThreadPlugin` so the
 * host's worker owns compute. Verified via `/sdk-harness?formulaWorker=1`. Own
 * spec to avoid colliding with sdk-harness.spec.ts.
 */

test('formula computes through an off-main worker', async ({ page }) => {
  await page.goto('/sdk-harness?formulaWorker=1');
  await page.waitForFunction(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    () => (window as any).__sdkHarnessReady === true,
    null,
    { timeout: 30_000 },
  );
  const result = await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__sdkHarnessAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.univer.getActiveWorkbook().getActiveSheet();
    // The RPC worker handshake completes ~1-2 s after onReady; a formula set
    // before the channel is up won't sync to the worker. Re-set on each poll so
    // the test is robust to handshake timing (a real user edits well after).
    for (let i = 0; i < 150; i++) {
      ws.getRange(0, 0).setValue({ f: '=1+2' });
      await new Promise((r) => setTimeout(r, 200));
      const v = ws.getRange(0, 0).getValue();
      if (v === 3 || v === '3') return v;
    }
    return ws.getRange(0, 0).getValue();
  });
  expect(Number(result)).toBe(3);
});
