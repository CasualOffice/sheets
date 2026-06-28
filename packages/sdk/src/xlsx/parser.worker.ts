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

import { workbookFromExcelJs } from './parse-impl';
import type { ImportedWorkbook } from './import';

/**
 * Worker side of the xlsx parser. Receives an ArrayBuffer via the
 * transferable list (zero-copy in), runs the pure ExcelJS conversion,
 * sends back the IWorkbookData. Errors are stringified — Workbook
 * objects don't structured-clone, but plain JSON does.
 *
 * The worker is single-shot per parse: parse-in-worker.ts spawns,
 * posts, awaits, terminates. Keeping it stateless side-steps any
 * worker-lifecycle bookkeeping for now; if we end up parsing many
 * small workbooks in succession we can switch to a pooled model.
 */

type ParseRequest = { id: number; buffer: ArrayBuffer };
type ParseResponse =
  | { id: number; ok: true; data: ImportedWorkbook }
  | { id: number; ok: false; error: string };

const ctx = self as unknown as DedicatedWorkerGlobalScope;

ctx.addEventListener('message', async (e: MessageEvent<ParseRequest>) => {
  const { id, buffer } = e.data;
  try {
    const data = await workbookFromExcelJs(buffer);
    const reply: ParseResponse = { id, ok: true, data };
    ctx.postMessage(reply);
  } catch (err) {
    // Preserve the underlying error type + stack so the main-thread
    // overlay can show something more useful than "Error code: 5".
    const error =
      err instanceof Error ? `${err.name}: ${err.message}\n${err.stack ?? ''}` : String(err);
    const reply: ParseResponse = { id, ok: false, error };
    ctx.postMessage(reply);
  }
});
