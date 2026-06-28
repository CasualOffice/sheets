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

/**
 * Round-trip helpers for the Watch Window list. Mirrors the pivots/charts
 * pattern: the watches live on `IWorkbookData.resources` and travel through
 * xlsx via the hidden `__casual_sheets_resources__` sidecar sheet, so a
 * Save → Open restores them. (Like pivots/charts, autosave's `wb.save()`
 * snapshot doesn't carry React-context state — that gap is shared.)
 */

import type { IWorkbookData } from '@univerjs/core';
import type { Watch } from './watch-model';

export const WATCHES_RESOURCE_NAME = '__casual_sheets_watches__';

type WatchesResourceV1 = { v: 1; watches: Watch[] };

function isValidWatch(w: unknown): w is Watch {
  if (!w || typeof w !== 'object') return false;
  const r = w as Record<string, unknown>;
  return (
    typeof r.id === 'string' &&
    typeof r.sheetId === 'string' &&
    typeof r.sheetName === 'string' &&
    typeof r.row === 'number' &&
    typeof r.col === 'number'
  );
}

/** Read watches off a snapshot. Tolerant of older / missing / malformed. */
export function readWatchesFromSnapshot(data: IWorkbookData | undefined): Watch[] {
  const entry = data?.resources?.find((r) => r.name === WATCHES_RESOURCE_NAME);
  if (!entry?.data) return [];
  try {
    const parsed = JSON.parse(entry.data) as Partial<WatchesResourceV1>;
    if (parsed?.v !== 1 || !Array.isArray(parsed.watches)) return [];
    return parsed.watches.filter(isValidWatch);
  } catch {
    return [];
  }
}

/** Merge watches INTO `data.resources` for export. Mutates in place. */
export function writeWatchesIntoSnapshot(data: IWorkbookData, watches: Watch[]): void {
  const filtered = (data.resources ?? []).filter((r) => r.name !== WATCHES_RESOURCE_NAME);
  if (watches.length === 0) {
    data.resources = filtered;
    return;
  }
  const payload: WatchesResourceV1 = { v: 1, watches };
  data.resources = [...filtered, { name: WATCHES_RESOURCE_NAME, data: JSON.stringify(payload) }];
}
