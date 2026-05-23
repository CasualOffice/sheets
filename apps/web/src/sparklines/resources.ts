import type { IWorkbookData } from '@univerjs/core';
import {
  SPARKLINES_RESOURCE_NAME,
  type SparklineModel,
  type SparklineType,
  type SparklinesResourceV1,
} from './types';

/**
 * Round-trip helpers for the sparklines plugin resource. Mirrors the
 * pattern used by `pivots/resources.ts` and `charts/resources.ts`:
 * the resource lives on `IWorkbookData.resources` and travels through
 * xlsx via the hidden `__casual_sheets_resources__` sheet, and
 * through collab via the snapshot-load path.
 *
 * Validation is defensive — older payloads, hand-edited JSON, or a
 * future-version schema all degrade to "no sparklines" rather than
 * crashing the workbook open.
 */

const VALID_TYPES: SparklineType[] = ['line', 'column', 'win-loss'];

function isValidSparkline(s: unknown): s is SparklineModel {
  if (!s || typeof s !== 'object') return false;
  const r = s as Record<string, unknown>;
  if (typeof r.id !== 'string') return false;
  if (!VALID_TYPES.includes(r.type as SparklineType)) return false;
  if (typeof r.unitId !== 'string' || typeof r.sheetId !== 'string') return false;
  const src = r.source as Record<string, unknown> | undefined;
  if (!src) return false;
  for (const k of ['startRow', 'endRow', 'startColumn', 'endColumn'] as const) {
    if (typeof src[k] !== 'number') return false;
  }
  const a = r.anchor as Record<string, unknown> | undefined;
  if (!a || typeof a.row !== 'number' || typeof a.col !== 'number') return false;
  return true;
}

export function readSparklinesFromSnapshot(data: IWorkbookData | undefined): SparklineModel[] {
  if (!data?.resources?.length) return [];
  const entry = data.resources.find((r) => r.name === SPARKLINES_RESOURCE_NAME);
  if (!entry?.data) return [];
  try {
    const parsed = JSON.parse(entry.data) as Partial<SparklinesResourceV1>;
    if (parsed?.v !== 1 || !Array.isArray(parsed.sparklines)) return [];
    return parsed.sparklines.filter(isValidSparkline);
  } catch {
    return [];
  }
}

export function writeSparklinesIntoSnapshot(data: IWorkbookData, sparklines: SparklineModel[]): void {
  const existing = data.resources ?? [];
  const filtered = existing.filter((r) => r.name !== SPARKLINES_RESOURCE_NAME);
  if (sparklines.length === 0) {
    data.resources = filtered;
    return;
  }
  const payload: SparklinesResourceV1 = { v: 1, sparklines };
  data.resources = [...filtered, { name: SPARKLINES_RESOURCE_NAME, data: JSON.stringify(payload) }];
}
