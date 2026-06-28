import type { IWorkbookData } from '@univerjs/core';
import {
  PIVOTS_RESOURCE_NAME,
  type PivotAggregation,
  type PivotModel,
  type PivotsResourceV1,
} from './types';

/**
 * Round-trip helpers for the pivots plugin resource. Mirrors the
 * pattern used by `charts/resources.ts` and `outline/resources.ts`:
 * the resource lives on `IWorkbookData.resources` and travels through
 * xlsx via the hidden `__casual_sheets_resources__` sheet, and
 * through collab via the snapshot-load path.
 */

const VALID_AGGS: PivotAggregation[] = [
  'sum',
  'count',
  'average',
  'min',
  'max',
  // Distinct Count (#212) — was missing here, so a saved Distinct-Count
  // pivot failed validation and was silently dropped on reload.
  'distinctCount',
];

function isValidPivot(p: unknown): p is PivotModel {
  if (!p || typeof p !== 'object') return false;
  const r = p as Record<string, unknown>;
  if (typeof r.id !== 'string') return false;
  if (typeof r.sourceSheetId !== 'string' || typeof r.targetSheetId !== 'string') return false;
  const src = r.source as Record<string, unknown> | undefined;
  if (!src) return false;
  for (const k of ['startRow', 'endRow', 'startColumn', 'endColumn'] as const) {
    if (typeof src[k] !== 'number') return false;
  }
  const tgt = r.target as Record<string, unknown> | undefined;
  if (!tgt || typeof tgt.row !== 'number' || typeof tgt.column !== 'number') return false;
  if (!Array.isArray(r.rows) || !Array.isArray(r.cols) || !Array.isArray(r.values)) return false;
  for (const f of r.rows as unknown[]) {
    if (!f || typeof f !== 'object' || typeof (f as { column?: unknown }).column !== 'number')
      return false;
  }
  for (const v of r.values as unknown[]) {
    if (!v || typeof v !== 'object') return false;
    const vv = v as { column?: unknown; agg?: unknown };
    if (typeof vv.column !== 'number') return false;
    if (!VALID_AGGS.includes(vv.agg as PivotAggregation)) return false;
  }
  return true;
}

/** Read pivot models out of a snapshot. Tolerant of older / missing payloads. */
export function readPivotsFromSnapshot(data: IWorkbookData | undefined): PivotModel[] {
  if (!data?.resources?.length) return [];
  const entry = data.resources.find((r) => r.name === PIVOTS_RESOURCE_NAME);
  if (!entry?.data) return [];
  try {
    const parsed = JSON.parse(entry.data) as Partial<PivotsResourceV1>;
    if (parsed?.v !== 1 || !Array.isArray(parsed.pivots)) return [];
    return parsed.pivots.filter(isValidPivot);
  } catch {
    /* corrupt payload — drop silently, the workbook still opens fine */
    return [];
  }
}

/** Merge pivot models INTO `data.resources` for export. Mutates in place. */
export function writePivotsIntoSnapshot(data: IWorkbookData, pivots: PivotModel[]): void {
  const existing = data.resources ?? [];
  const filtered = existing.filter((r) => r.name !== PIVOTS_RESOURCE_NAME);
  if (pivots.length === 0) {
    data.resources = filtered;
    return;
  }
  const payload: PivotsResourceV1 = { v: 1, pivots };
  data.resources = [...filtered, { name: PIVOTS_RESOURCE_NAME, data: JSON.stringify(payload) }];
}
