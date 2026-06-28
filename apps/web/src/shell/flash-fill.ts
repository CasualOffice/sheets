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

import type { FUniver } from '@univerjs/core/facade';
import type { FWorksheet } from '@univerjs/sheets/facade';
import { rangeAt, sheetId as facadeSheetId } from '../univer-facade';

/**
 * Excel's Ctrl+E "Flash Fill" — heuristic pattern detection.
 *
 * Given a target column where the user has typed a handful of examples
 * derived from the column(s) to the left, infer the transform and
 * apply it to the empty rows below. v1 handles the patterns Flash
 * Fill demos most often:
 *
 *   - **Identity** — `John Doe → John Doe`
 *   - **Token-at-index** — `John Doe → John` (split by space, [0])
 *   - **Token-at-index** — `john@acme.com → acme.com` (split by @, [1])
 *   - **Substring** — `2024-Q3-report → Q3` (chars 5..7)
 *   - **Case** — `JOHN DOE → John Doe` (lowercase + title)
 *
 * Multi-source-column concat patterns are intentionally out of scope
 * for v1 — they account for <10% of demoed Flash Fill use cases and
 * would double the algorithm's complexity. Easy follow-up if it turns
 * out users want it.
 *
 * All writes go through `set-range-values`, so co-edit propagation and
 * xlsx round-trip are automatic — no special handling needed.
 */

type Cell = string;

/**
 * Candidate transform — pure function from a source string to a
 * candidate fill. Generated from each (source, example) pair; we keep
 * the one that satisfies every pair.
 */
type Transform = (src: string) => string;

const COMMON_SEPARATORS = [' ', '@', ',', '.', '-', '_', '/', ':', ';', '|', '\t'];

/**
 * Try to derive a transform from (source, example). Returns null if
 * no obvious pattern matches.
 */
function inferTransform(source: string, example: string): Transform | null {
  if (source === example) return (s) => s;

  // Whole-source case transforms.
  if (example === source.toLowerCase()) return (s) => s.toLowerCase();
  if (example === source.toUpperCase()) return (s) => s.toUpperCase();
  if (example === toTitleCase(source)) return (s) => toTitleCase(s);

  // Token-at-index (try each separator + each index).
  for (const sep of COMMON_SEPARATORS) {
    const parts = source.split(sep);
    if (parts.length <= 1) continue;
    for (let i = 0; i < parts.length; i += 1) {
      if (parts[i] === example) {
        return (s) => s.split(sep)[i] ?? '';
      }
      // Joined tail / head, e.g. take everything after the first @.
      const tail = parts.slice(i).join(sep);
      if (tail === example) {
        return (s) => s.split(sep).slice(i).join(sep);
      }
      const head = parts.slice(0, i + 1).join(sep);
      if (head === example) {
        return (s) => s.split(sep).slice(0, i + 1).join(sep);
      }
    }
  }

  // Plain substring — find example inside source.
  const idx = source.indexOf(example);
  if (idx !== -1 && example.length > 0) {
    // Anchor by position (start) and length. Brittle but matches Excel's
    // simplest substring inference for short examples.
    const start = idx;
    const len = example.length;
    return (s) => s.substring(start, start + len);
  }

  // Prefix + source.
  if (example.endsWith(source)) {
    const prefix = example.slice(0, example.length - source.length);
    return (s) => prefix + s;
  }
  // Source + suffix.
  if (example.startsWith(source)) {
    const suffix = example.slice(source.length);
    return (s) => s + suffix;
  }

  return null;
}

function toTitleCase(s: string): string {
  return s.replace(
    /\w\S*/g,
    (w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase(),
  );
}

/**
 * From a list of (source, example) pairs, find a transform that
 * satisfies ALL pairs. We try each pair's candidate transform in turn
 * and accept the first one that works on every other pair.
 */
function deriveSharedTransform(
  pairs: Array<{ source: Cell; example: Cell }>,
): Transform | null {
  for (const pair of pairs) {
    const candidate = inferTransform(pair.source, pair.example);
    if (!candidate) continue;
    const fitsAll = pairs.every((p) => candidate(p.source) === p.example);
    if (fitsAll) return candidate;
  }
  return null;
}

/**
 * Run Flash Fill from the active selection. Walks LEFT from the
 * selection's column to find the nearest column with values (the
 * "source"), reads the example values the user typed in the target
 * column, infers the transform, and writes the fill below.
 *
 * Returns one of four outcomes; `'filled'` carries the count for
 * UI confirmation.
 *   - `{ status: 'filled', count }` — wrote `count` cells
 *   - `{ status: 'no-pattern' }` — couldn't find a transform
 *   - `{ status: 'no-source' }` — no source column to derive from
 *   - `{ status: 'no-examples' }` — target column has nothing to learn from
 */
export type FlashFillResult =
  | { status: 'filled'; count: number }
  | { status: 'no-pattern' }
  | { status: 'no-source' }
  | { status: 'no-examples' };

export function flashFill(api: FUniver): FlashFillResult {
  const wb = api.getActiveWorkbook();
  const sheet = wb?.getActiveSheet() ?? null;
  if (!wb || !sheet) return { status: 'no-source' };
  const range = sheet.getActiveRange();
  if (!range) return { status: 'no-examples' };

  const startRow: number = range.getRow();
  const startCol: number = range.getColumn();
  const height: number = range.getHeight();
  const width: number = range.getWidth();

  if (width !== 1) {
    // Multi-column selections aren't a Flash-Fill use case; bail.
    return { status: 'no-examples' };
  }

  // Find source column: nearest non-empty column to the LEFT.
  let sourceCol = -1;
  for (let c = startCol - 1; c >= 0; c -= 1) {
    let hasValue = false;
    for (let r = startRow; r < startRow + height; r += 1) {
      const v = readCell(sheet, r, c);
      if (v && v.length > 0) {
        hasValue = true;
        break;
      }
    }
    if (hasValue) {
      sourceCol = c;
      break;
    }
  }
  if (sourceCol === -1) return { status: 'no-source' };

  // Read all (source, target) pairs in the selection rows.
  const examples: Array<{ source: Cell; example: Cell }> = [];
  const blanks: Array<{ row: number; source: Cell }> = [];
  for (let r = startRow; r < startRow + height; r += 1) {
    const source = readCell(sheet, r, sourceCol);
    const target = readCell(sheet, r, startCol);
    if (!source) continue;
    if (target && target.length > 0) {
      examples.push({ source, example: target });
    } else {
      blanks.push({ row: r, source });
    }
  }

  if (examples.length === 0) return { status: 'no-examples' };
  if (blanks.length === 0) return { status: 'no-examples' };

  const transform = deriveSharedTransform(examples);
  if (!transform) return { status: 'no-pattern' };

  // Apply the transform to every blank row in the selection.
  const filled: Array<{ row: number; value: string }> = blanks.map((b) => ({
    row: b.row,
    value: transform(b.source),
  }));

  const unitId = wb.getId();
  const subUnitId = facadeSheetId(sheet);
  // Single set-range-values command containing every fill — one
  // undo entry, one collab broadcast. The command takes a `value`
  // param shaped as the object-matrix Univer expects (row → col → cell).
  const value: Record<number, Record<number, { v: string }>> = {};
  for (const f of filled) {
    value[f.row] = { [startCol]: { v: f.value } };
  }
  void api.executeCommand('sheet.command.set-range-values', {
    unitId,
    subUnitId,
    value,
  });

  return { status: 'filled', count: filled.length };
}

function readCell(sheet: FWorksheet, row: number, col: number): string {
  try {
    const range = rangeAt(sheet, row, col);
    const data = range?.getCellData();
    if (data == null) return '';
    const v = data.v;
    if (v == null) return '';
    return String(v);
  } catch {
    return '';
  }
}
