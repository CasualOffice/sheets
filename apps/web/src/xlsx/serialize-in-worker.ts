import type { IWorkbookData } from '@univerjs/core';
import type { ExportExtras } from './export';

/**
 * Main-thread entry point for xlsx export. Spawns a Web Worker, hands
 * over the snapshot via structured clone (Yjs awareness updates can
 * happen in parallel — the worker gets a frozen copy), waits for the
 * resulting Blob, terminates.
 *
 * Why structured-cloned rather than transferable: the snapshot is a
 * plain JS tree, and transferables would detach it from the main thread,
 * breaking the UI that's still rendering it.
 */

let nextId = 0;

export function serializeXlsxInWorker(
  data: IWorkbookData,
  extras: ExportExtras = {},
): Promise<Blob> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL('./exporter.worker.ts', import.meta.url), {
      type: 'module',
      name: 'xlsx-exporter',
    });
    const id = ++nextId;
    const cleanup = () => worker.terminate();
    worker.addEventListener(
      'message',
      (e: MessageEvent<{ id: number; ok: true; blob: Blob } | { id: number; ok: false; error: string }>) => {
        const msg = e.data;
        if (msg.id !== id) return;
        cleanup();
        if (msg.ok) resolve(msg.blob);
        else reject(new Error(msg.error));
      },
    );
    worker.addEventListener('error', (e) => {
      cleanup();
      reject(new Error(`xlsx exporter error: ${e.message || 'unknown'}`));
    });
    worker.postMessage({ id, data, extras });
  });
}
