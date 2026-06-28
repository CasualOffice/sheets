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
