import { expect, test, chromium, type Browser } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Regression for the divergence indicator added in feat(collab):
 * peers broadcast their Y.Doc state-vector hex via awareness and the
 * CollabIndicator flips to amber `data-sync-health=diverged` when the
 * local SV disagrees with every peer's SV for >15 s.
 *
 * Forcing the disagreement: write a fake `sv` into the joiner's
 * awareness state directly, then leave it. The owner's CollabIndicator
 * should transition `in-sync` → `syncing` → `diverged` over ~17 s.
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

test('divergence indicator flips amber when peer state-vectors disagree', async () => {
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

  // Wait for the initial in-sync state to be reported by both pills.
  await expect(owner.getByTestId('collab-indicator')).toHaveAttribute(
    'data-sync-health',
    'in-sync',
    { timeout: 10_000 },
  );

  // Force divergence by monkey-patching the joiner's
  // awareness.setLocalState so EVERY write — including the bridge's
  // periodic real-SV broadcasts — is rewritten to a bogus SV. We
  // can't just write once because usePresenceWire's heartbeat
  // overwrites it; we can't race a 1s setInterval against the
  // bridge's 5s heartbeat reliably either. Patching the setter wins
  // every time.
  await joiner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const provider = (window as any).__hocuspocusProvider;
    if (!provider?.awareness) throw new Error('awareness not exposed');
    const aw = provider.awareness;
    const orig = aw.setLocalState.bind(aw);
    const bad = 'deadbeefdeadbeefdeadbeef'; // arbitrary fake SV
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    aw.setLocalState = (s: any) => {
      orig({ ...(s ?? {}), sv: bad, svAt: Date.now() });
    };
    // Trigger one immediate write so the next observer fire on the
    // owner's side sees our fake.
    aw.setLocalState(aw.getLocalState() ?? {});
  });

  // After the 15 s grace window plus a 2 s recompute interval, the
  // owner's indicator must read `diverged`.
  await expect(owner.getByTestId('collab-indicator')).toHaveAttribute(
    'data-sync-health',
    'diverged',
    { timeout: 25_000 },
  );

  // UI also shows the amber pill text.
  await expect(owner.getByTestId('collab-indicator')).toContainText('Out of sync');

  // Clean up the divergence forcer so the test session doesn't leave
  // residual intervals on the page.
  await joiner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const id = (window as any).__divergeInterval;
    if (id) clearInterval(id);
  });

  await ownerCtx.close();
  await joinerCtx.close();
});
