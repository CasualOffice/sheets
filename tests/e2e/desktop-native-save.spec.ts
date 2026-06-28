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

import { expect, test } from '@playwright/test';
import ExcelJS from 'exceljs';
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

/**
 * Mid-save staleness guard. The bridge marks the window clean after a save ONLY
 * if the document hasn't changed since the bytes were serialized — the caller
 * pins the edit counter at serialization (`save(bytes, baselineSeq)`). An edit
 * that lands between serialize and the disk write must keep the window dirty so
 * it isn't silently lost on close. Drives the REAL bridge (mocking only the
 * Tauri invoke layer), not a fake bridge, so it exercises that decision.
 */
test('desktop bridge stays dirty when an edit lands between serialize and write', async ({
  page,
}) => {
  test.setTimeout(60_000);

  await page.addInitScript(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rec: any = { setDirty: [] as boolean[] };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__deskTest = rec;
    // Minimal Tauri invoke mock: record set_window_dirty transitions; resolve
    // the chunked-write commands so chunkedWrite completes without a real FS.
    const invoke = (cmd: string, args?: { dirty?: boolean }) => {
      if (cmd === 'set_window_dirty') rec.setDirty.push(args!.dirty);
      if (cmd === 'document_size') return Promise.resolve(0);
      return Promise.resolve(null);
    };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__TAURI__ = { core: { invoke } };
  });

  // ?desk=1 (top-level) activates the real bridge bootstrap; no file= so the
  // app doesn't try to auto-load a path — we bind filePath directly below.
  await page.goto('/?desk=1');
  await page.waitForFunction(() => !!(window as { __deskApp__?: unknown }).__deskApp__, null, {
    timeout: 30_000,
  });

  const out = await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const b: any = (window as any).__deskApp__;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rec: any = (window as any).__deskTest;
    b.filePath = '/tmp/book.xlsx';
    const bytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04]).buffer; // non-empty (PK header)

    // Stale save: capture the baseline, THEN an edit lands, THEN write.
    const staleBaseline = b.currentEditSeq();
    b.setDirty(true); // editSeq bumps → now ahead of staleBaseline
    await b.save(bytes, staleBaseline);
    const afterStale = [...rec.setDirty];

    // Clean save: baseline reflects the current edit state.
    await b.save(bytes, b.currentEditSeq());
    return { afterStale, afterClean: [...rec.setDirty] };
  });

  // The dirty edit fired set_window_dirty(true); the stale save must NOT clear it.
  expect(out.afterStale).toEqual([true]);
  // The clean save (baseline === current) clears it → a false transition appended.
  expect(out.afterClean).toEqual([true, false]);
});

/**
 * Concurrent-write serialization. Two overlapping saves — a fast double Ctrl+S,
 * or Ctrl+S while a large save is still streaming — must NOT interleave their
 * begin/write/commit IPC against the shared per-path temp file (that would
 * corrupt it). The bridge chains writes so each completes before the next
 * starts. Drives the REAL bridge with a deliberately slow write_save_chunk to
 * force overlap, then asserts the recorded IPC order never interleaves.
 */
test('desktop bridge serializes overlapping saves (no interleaved write IPC)', async ({ page }) => {
  test.setTimeout(60_000);

  await page.addInitScript(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rec: any = { order: [] as string[] };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__deskTest = rec;
    const delay = (ms: number) => new Promise((r) => setTimeout(r, ms));
    const invoke = async (cmd: string) => {
      if (cmd === 'document_size') return 0;
      if (
        cmd === 'begin_save_document' ||
        cmd === 'write_save_chunk' ||
        cmd === 'commit_save_document'
      ) {
        rec.order.push(cmd);
        // Slow the chunk write so a second save would overlap if unserialized.
        if (cmd === 'write_save_chunk') await delay(60);
      }
      return null;
    };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__TAURI__ = { core: { invoke } };
  });

  await page.goto('/?desk=1');
  await page.waitForFunction(() => !!(window as { __deskApp__?: unknown }).__deskApp__, null, {
    timeout: 30_000,
  });

  const order = await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const b: any = (window as any).__deskApp__;
    b.filePath = '/tmp/book.xlsx';
    const bytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04]).buffer; // non-empty
    // Fire two saves without awaiting the first — they race unless serialized.
    await Promise.all([b.save(bytes), b.save(bytes)]);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (window as any).__deskTest.order as string[];
  });

  // Two full, non-interleaved cycles back to back.
  expect(order).toEqual([
    'begin_save_document',
    'write_save_chunk',
    'commit_save_document',
    'begin_save_document',
    'write_save_chunk',
    'commit_save_document',
  ]);
});

/**
 * Short-read guard on open. The chunked read sizes its buffer to
 * `document_size`; if the file shrinks/changes mid-read (another process
 * truncates or replaces it — the external edits the file watcher reports), a
 * chunk comes back short and the tail would be zero-padded, parsing as a
 * baffling corruption. The bridge must instead fail with a clear "file changed
 * while opening" error so the caller can re-open. Drives the REAL bridge with a
 * truncated read_document_chunk.
 */
test('desktop bridge fails clearly when a file is truncated mid-read', async ({ page }) => {
  test.setTimeout(60_000);

  await page.addInitScript(() => {
    // document_size says 10 bytes, but the read yields only 4 then 0 — a file
    // that shrank between sizing and reading.
    let reads = 0;
    const invoke = (cmd: string) => {
      if (cmd === 'document_size') return Promise.resolve(10);
      if (cmd === 'read_document_chunk') {
        reads += 1;
        return Promise.resolve(reads === 1 ? [0x50, 0x4b, 0x03, 0x04] : []);
      }
      return Promise.resolve(null);
    };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__TAURI__ = { core: { invoke } };
  });

  await page.goto('/?desk=1');
  await page.waitForFunction(() => !!(window as { __deskApp__?: unknown }).__deskApp__, null, {
    timeout: 30_000,
  });

  const err = await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const b: any = (window as any).__deskApp__;
    try {
      await b.loadDocument('/tmp/book.xlsx');
      return null;
    } catch (e) {
      return (e as Error).message;
    }
  });

  expect(err).toContain('the file changed while opening');
  // It must NOT have returned a zero-padded buffer (4 real bytes of 10).
  expect(err).toContain('Only read 4 of 10 bytes');
});

/**
 * Reload retry on a transient mid-write. When another app saves the open file,
 * the watcher often fires while the write is still in flight, so the first
 * reload reads short (the guard above throws). Rather than leaving the user on
 * stale content, the reload retries once after a short settle delay — by then
 * the external write has completed and the reload succeeds. Drives the real app
 * + bridge: first read is truncated, the retry serves a valid workbook.
 */
test('desktop reload retries once when an external save is caught mid-write', async ({ page }) => {
  test.setTimeout(60_000);

  // The valid workbook served on the successful retry (A1 = "RELOADED").
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'RELOADED';
  const xlsxBytes = Array.from(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));

  await page.addInitScript((bytes: number[]) => {
    let loadCount = 0; // one per loadDocument (document_size kicks each off)
    let reads = 0;
    const invoke = (cmd: string) => {
      if (cmd === 'document_size') {
        loadCount += 1;
        reads = 0;
        return Promise.resolve(bytes.length);
      }
      if (cmd === 'read_document_chunk') {
        reads += 1;
        // First reload attempt: a short read then empty → short-read throw.
        if (loadCount === 1) return Promise.resolve(reads === 1 ? bytes.slice(0, 4) : []);
        // Retry attempt: serve the whole file.
        return Promise.resolve(reads === 1 ? bytes : []);
      }
      return Promise.resolve(null);
    };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__TAURI__ = { core: { invoke } };
  }, xlsxBytes);

  await page.goto('/?desk=1');
  await waitForUniver(page);

  // Bind a path and fire the external-change event the file watcher would emit.
  await page.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__deskApp__.filePath = '/tmp/x.xlsx';
    window.dispatchEvent(
      new CustomEvent('deskapp:file-changed', {
        detail: { kind: 'modified', path: '/tmp/x.xlsx' },
      }),
    );
  });

  // First reload throws (short read); the retry serves the full file and A1
  // becomes "RELOADED". Without the retry, A1 would stay empty and this fails.
  await expect
    .poll(
      () =>
        page.evaluate(() => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const api = (window as any).__univerAPI;
          try {
            return api?.getActiveWorkbook()?.getActiveSheet()?.getRange('A1')?.getValue() ?? null;
          } catch {
            return null;
          }
        }),
      { timeout: 10_000 },
    )
    .toBe('RELOADED');
});
