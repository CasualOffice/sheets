import type { ImportedWorkbook } from './import';

/**
 * Main-thread entry point for xlsx parsing. Spawns a Web Worker,
 * transfers the buffer in (zero-copy), awaits a single response,
 * terminates the worker. The buffer becomes detached on the main
 * side after transfer — callers must not reuse it.
 *
 * Why a fresh worker per parse: parses are rare (open / drop /
 * collab seed) and the worker carries a ~600 KB ExcelJS payload —
 * keeping it alive between parses would idle that memory. A pool
 * is a follow-up optimization if we see many sequential opens.
 */

let nextId = 0;

// Cap the wait at 3 minutes. A multi-hundred-MB ExcelJS parse can take
// over a minute on slow hardware; anything past 3 minutes is almost
// certainly a hung/deadlocked worker (e.g. crashed without firing
// `error`). Reject so the UI shows an actionable message instead of
// spinning forever.
const PARSE_TIMEOUT_MS = 180_000;

export function parseXlsxInWorker(buffer: ArrayBuffer): Promise<ImportedWorkbook> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL('./parser.worker.ts', import.meta.url), {
      type: 'module',
      name: 'xlsx-parser',
    });
    const id = ++nextId;
    let timeoutHandle: ReturnType<typeof setTimeout> | null = null;
    const cleanup = () => {
      if (timeoutHandle !== null) clearTimeout(timeoutHandle);
      worker.terminate();
    };
    timeoutHandle = setTimeout(() => {
      cleanup();
      reject(
        new Error(
          `xlsx parser worker timed out after ${Math.round(PARSE_TIMEOUT_MS / 1000)}s. The workbook may be corrupted or too large to parse in this browser.`,
        ),
      );
    }, PARSE_TIMEOUT_MS);
    worker.addEventListener(
      'message',
      (e: MessageEvent<{ id: number; ok: true; data: ImportedWorkbook } | { id: number; ok: false; error: string }>) => {
        const msg = e.data;
        if (msg.id !== id) return;
        cleanup();
        if (msg.ok) resolve(msg.data);
        else reject(new Error(msg.error));
      },
    );
    worker.addEventListener('error', (e) => {
      cleanup();
      // ErrorEvent.message is "" when the worker crashes from OOM in
      // some browsers; with no message AND no filename it's almost
      // always OOM mid-allocation. Emit a clear hint instead of the
      // useless "(unknown)" we used to show.
      const hasDetail = Boolean(e.message || e.filename);
      if (!hasDetail) {
        reject(
          new Error(
            'xlsx parser worker ran out of memory parsing this file. Try a smaller workbook or close other browser tabs.',
          ),
        );
        return;
      }
      const detail = e.message || `${e.filename ?? ''}:${e.lineno ?? ''}`;
      reject(
        new Error(
          `xlsx parser worker crashed (${detail}). The workbook may be too large for this browser to allocate.`,
        ),
      );
    });
    worker.addEventListener('messageerror', () => {
      cleanup();
      reject(new Error('xlsx parser worker returned an unserializable result.'));
    });
    worker.postMessage({ id, buffer }, [buffer]);
  });
}
