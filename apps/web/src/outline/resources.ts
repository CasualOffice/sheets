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

import type { IWorkbookData } from '@univerjs/core';
import { OUTLINE_RESOURCE_NAME, type OutlineResourceV1, type OutlineState } from './types';

/**
 * Round-trip helpers for the outline plugin resource. The resource lives on
 * `IWorkbookData.resources` and is carried through xlsx via the hidden
 * __casual_sheets_resources__ sheet we already use for other plugin state.
 */

/** Read outline state out of a snapshot. Tolerant of older / missing payloads. */
export function readOutlineFromSnapshot(data: IWorkbookData | undefined): OutlineState {
  if (!data?.resources?.length) return {};
  const entry = data.resources.find((r) => r.name === OUTLINE_RESOURCE_NAME);
  if (!entry?.data) return {};
  try {
    const parsed = JSON.parse(entry.data) as Partial<OutlineResourceV1>;
    if (parsed?.v !== 1 || !parsed.sheets || typeof parsed.sheets !== 'object') return {};
    return parsed.sheets;
  } catch {
    /* corrupt payload — drop silently, the workbook still opens fine */
    return {};
  }
}

/** Merge outline state INTO `data.resources` for export. Mutates in place. */
export function writeOutlineIntoSnapshot(
  data: IWorkbookData,
  state: OutlineState,
): void {
  const payload: OutlineResourceV1 = { v: 1, sheets: state };
  const serialized = JSON.stringify(payload);
  const existing = data.resources ?? [];
  const filtered = existing.filter((r) => r.name !== OUTLINE_RESOURCE_NAME);
  // Only write the resource if there's actually any group to persist —
  // empty state has no business polluting the resources array.
  const hasGroups = Object.values(state).some(
    (s) => s.rows.length > 0 || s.cols.length > 0,
  );
  data.resources = hasGroups
    ? [...filtered, { name: OUTLINE_RESOURCE_NAME, data: serialized }]
    : filtered;
}
