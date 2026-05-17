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

export function parseXlsxInWorker(buffer: ArrayBuffer): Promise<ImportedWorkbook> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL('./parser.worker.ts', import.meta.url), {
      type: 'module',
      name: 'xlsx-parser',
    });
    const id = ++nextId;
    const cleanup = () => worker.terminate();
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
      // some browsers; fall back to filename/line for context. Common
      // cause: a workbook too large to allocate during ExcelJS load.
      const detail = e.message || `${e.filename ?? ''}:${e.lineno ?? ''}` || 'unknown';
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
