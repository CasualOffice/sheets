import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Native save routing — Phase, desktop slice 3.
 *
 * In the desktop (Tauri) build, Save must overwrite the bound file on disk via
 * the `window.__deskApp__` bridge, never produce a browser download (the desktop
 * shell's hard rule). `deliverBlob` routes through the bridge whenever
 * `__deskApp__.isDesktop` is set. Here we inject a mock bridge and assert
 * Ctrl+S reaches `bridge.save()` with the workbook bytes.
 */
test('Ctrl+S routes through the desktop bridge instead of downloading', async ({ page }) => {
  test.setTimeout(60_000);

  // Inject a fake desktop bridge before the app boots. Records save() calls.
  await page.addInitScript(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__deskSave = { saveCalls: 0, saveAsCalls: 0, lastBytes: 0 };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__deskApp__ = {
      isDesktop: true,
      filePath: '/Users/test/book.xlsx',
      save(bytes: ArrayBuffer) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const rec = (window as any).__deskSave;
        rec.saveCalls += 1;
        rec.lastBytes = bytes.byteLength;
        return Promise.resolve('/Users/test/book.xlsx');
      },
      saveAs(name: string, _bytes: ArrayBuffer) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (window as any).__deskSave.saveAsCalls += 1;
        return Promise.resolve('/Users/test/' + name);
      },
    };
  });

  await page.goto('/');
  await waitForUniver(page);

  // Make an edit so there's something to save.
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (api.getActiveWorkbook()!.getActiveSheet() as any).getRange('A1').setValue({ v: 'desk save' });
  });

  await page.keyboard.press('Control+s');

  // Bridge.save() was called with non-empty bytes; no saveAs (path is bound).
  await expect
    .poll(() =>
      page.evaluate(() => (window as { __deskSave?: { saveCalls: number } }).__deskSave!.saveCalls),
    )
    .toBeGreaterThanOrEqual(1);
  const rec = await page.evaluate(
    () => (window as { __deskSave?: { saveAsCalls: number; lastBytes: number } }).__deskSave!,
  );
  expect(rec.saveAsCalls).toBe(0);
  expect(rec.lastBytes).toBeGreaterThan(0);
});

/**
 * Dirty signal for the native close-guard. The bridge's dirty bit is driven by
 * the editor's command-bus mutation hook (`onMutationExecutedForCollab`) — the
 * project's only sanctioned change hook — via `bridge.setDirty(true)`, NOT a
 * DOM-keystroke heuristic. This proves a non-keyboard edit (a programmatic
 * `setValue`, like paste / toolbar / fill would produce) still marks dirty.
 */
test('an edit drives bridge.setDirty(true) via the mutation hook', async ({ page }) => {
  test.setTimeout(60_000);

  await page.addInitScript(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__deskDirty = { calls: 0, last: null as boolean | null };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__deskApp__ = {
      isDesktop: true,
      filePath: '/Users/test/book.xlsx',
      save: (_b: ArrayBuffer) => Promise.resolve('/Users/test/book.xlsx'),
      saveAs: (n: string, _b: ArrayBuffer) => Promise.resolve('/Users/test/' + n),
      setDirty(dirty: boolean) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const rec = (window as any).__deskDirty;
        rec.calls += 1;
        rec.last = dirty;
      },
    };
  });

  await page.goto('/');
  await waitForUniver(page);

  // Programmatic edit — no keystrokes at all (the case the old DOM heuristic
  // missed). Goes through the command bus, so the mutation hook must fire.
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (api.getActiveWorkbook()!.getActiveSheet() as any).getRange('A1').setValue({ v: 'dirty' });
  });

  await expect
    .poll(() =>
      page.evaluate(() => (window as { __deskDirty?: { calls: number } }).__deskDirty!.calls),
    )
    .toBeGreaterThanOrEqual(1);
  const last = await page.evaluate(
    () => (window as { __deskDirty?: { last: boolean | null } }).__deskDirty!.last,
  );
  expect(last).toBe(true);
});
