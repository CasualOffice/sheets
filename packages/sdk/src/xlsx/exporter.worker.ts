/// <reference lib="webworker" />

import { workbookDataToXlsxImpl } from './export-impl';
import type { IWorkbookData } from '@univerjs/core';
import type { ExportExtras } from './export';

/**
 * Worker side of the xlsx exporter. Receives the snapshot + extras as
 * structured-cloned JSON, runs the ExcelJS serialization, posts back a Blob.
 * Single-shot per export — serialize-in-worker.ts spawns, posts, awaits,
 * terminates.
 */

type ExportRequest = { id: number; data: IWorkbookData; extras: ExportExtras };
type ExportResponse =
  | { id: number; ok: true; blob: Blob }
  | { id: number; ok: false; error: string };

const ctx = self as unknown as DedicatedWorkerGlobalScope;

ctx.addEventListener('message', async (e: MessageEvent<ExportRequest>) => {
  const { id, data, extras } = e.data;
  try {
    const blob = await workbookDataToXlsxImpl(data, extras);
    const reply: ExportResponse = { id, ok: true, blob };
    ctx.postMessage(reply);
  } catch (err) {
    const reply: ExportResponse = { id, ok: false, error: String(err) };
    ctx.postMessage(reply);
  }
});
