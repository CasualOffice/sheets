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

import { LocaleType, type IWorkbookData } from '@univerjs/core';
import appPkg from '../package.json';

// Sourced from our package.json's `@univerjs/core` dependency declaration so
// the snapshot's appVersion tracks the installed Univer release automatically.
// (Vite supports JSON imports; tsconfig has resolveJsonModule.)
const UNIVER_DEP = (appPkg.dependencies as Record<string, string>)['@univerjs/core'];
export const UNIVER_VERSION = UNIVER_DEP.replace(/^[~^]/, '');

/**
 * Initial workbook size. Univer materializes row/column metadata for the
 * declared count, so we keep this modest to boot fast. The grid grows
 * dynamically (see `useWorkbookGrowth`) up to MAX_ROWS / MAX_COLUMNS.
 */
export const INITIAL_ROWS = 1024;
// 26 = A..Z. Univer allocates row/column metadata up-front for the declared
// count, so a 128-wide start cost ~5× the boot allocation for columns the
// user almost never reaches before `useWorkbookGrowth` extends them.
export const INITIAL_COLUMNS = 26;
// Interactive ceiling — Excel parity (1,048,576 × 16,384). Safe because cost
// tracks usage, not these constants:
//   - `useWorkbookGrowth` extends rowCount/columnCount in small on-demand chunks;
//     small sheets stay small.
//   - The render skeleton's accumulation arrays (vendor/univer-revamp/.../
//     sheet-skeleton.ts) are only as large as the grown/imported extent, and a
//     pure cell edit does NOT rebuild them.
// Measured (chromium, M-series): declaring a full 1,048,576-row grid + editing
// the far edge is a one-time ~170 ms; a pure cell edit at 1M rows is ~1 ms. The
// only costly op is insert/delete-row at extreme row counts (~517 ms @1M), which
// is optimized separately by sparse insert/delete (docs/COMPETITIVE_ROADMAP.md
// Phase 2, T2.1 / UNIVER_FORK_PERF.md item 6) — not a regression of normal use.
export const MAX_ROWS = 1048576;
export const MAX_COLUMNS = 16384;

export function emptyWorkbook(): IWorkbookData {
  const nowIso = new Date().toISOString();
  return {
    // Unique per call — Univer's IUniverInstanceService rejects duplicate unit
    // ids, so a fresh blank workbook must not collide with the one it's
    // replacing.
    id: `wb-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
    rev: 1,
    name: 'Untitled',
    appVersion: UNIVER_VERSION,
    locale: LocaleType.EN_US,
    styles: {},
    sheetOrder: ['sheet-1'],
    sheets: {
      'sheet-1': {
        id: 'sheet-1',
        name: 'Sheet1',
        cellData: {},
        rowCount: INITIAL_ROWS,
        columnCount: INITIAL_COLUMNS,
      },
    },
    // Stamp creation time up front so a never-saved workbook still has
    // an honest "Created" in the Properties dialog. modifiedAt bumps
    // on every Save (export-impl always writes new Date() to xlsx),
    // and re-reads on the next open.
    custom: {
      properties: { createdAt: nowIso, modifiedAt: nowIso },
    },
  };
}
