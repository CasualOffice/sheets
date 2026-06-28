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

import { useEffect } from 'react';
import { ICommandService, type ICommandInfo, type IExecutionOptions } from '@univerjs/core';
import { useUniverAPI } from '../use-univer';
import { parseExcelStyleValue } from './excel-input-parser';

/**
 * Intercepts the cell-edit commit path so typed input that contains
 * Excel-style decorations (currency symbol, percent suffix, accounting
 * parentheses, thousands grouping) lands as a numeric cell instead of
 * text.
 *
 * Univer's in-cell editor commits via `sheet.command.set-range-values`
 * with the typed string as `v` and no `t`, which Univer treats as
 * STRING — that's what trips the "Number stored as text" alert when
 * the user types `$2,000` into a currency-formatted cell.
 *
 * We mutate `info.params.cellValue` (or `value`) in place during the
 * `beforeCommandExecuted` hook, BEFORE Univer applies the mutation, so
 * both the snapshot and the collab broadcast carry the corrected
 * number. Guarded with `fromCollab` to avoid re-parsing mutations
 * arriving from peers (those were already parsed on the producer).
 *
 * We deliberately do NOT change the cell's number-format pattern here —
 * the original symptom is "value stored as text"; coercing the value is
 * sufficient for the cell's existing format to render correctly. The
 * pattern-application enhancement (e.g. auto-applying `"$"#,##0.00`
 * when typing into a blank cell) is a follow-up.
 */
export function InputParserDriver() {
  const api = useUniverAPI();

  useEffect(() => {
    if (!api) return;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const injector = (api as any)._injector as { get: (token: unknown) => unknown } | undefined;
    if (!injector) return;
    const cmdSvc = injector.get(ICommandService) as {
      beforeCommandExecuted: (l: (info: ICommandInfo, options?: IExecutionOptions) => void) => {
        dispose: () => void;
      };
    };

    const sub = cmdSvc.beforeCommandExecuted((info, options) => {
      if (options?.fromCollab) return;
      if (info.id !== 'sheet.command.set-range-values') return;
      const params = info.params as Record<string, unknown> | undefined;
      if (!params) return;
      // Univer command schema for set-range-values has used both
      // `value` and `cellValue` across versions; check both.
      if ('value' in params) walkAndCoerce(params, 'value');
      if ('cellValue' in params) walkAndCoerce(params, 'cellValue');
    });

    return () => sub.dispose();
  }, [api]);

  return null;
}

/**
 * Mutate the cell-value tree at `params[key]` in place. Supports the
 * three shapes Univer passes:
 *   - single `ICellData`: `{ v: "..." }`
 *   - 2D array: `[[{ v: "..." }, ...], ...]`
 *   - sparse row-major: `{ [row]: { [col]: { v: "..." } } }`
 */
function walkAndCoerce(params: Record<string, unknown>, key: string): void {
  const tree = params[key];
  if (tree == null) return;

  if (Array.isArray(tree)) {
    for (const row of tree) {
      if (Array.isArray(row)) {
        for (const cell of row) coerceCell(cell);
      } else {
        coerceCell(row);
      }
    }
    return;
  }

  if (typeof tree === 'object') {
    // Heuristic: a leaf ICellData has at least one of `v`, `f`, `t`,
    // `p`, `s`, `custom` as keys. A row-major map has rows-as-keys
    // whose values are themselves objects. We try the leaf shape
    // first; if it doesn't apply, we walk.
    if (looksLikeCell(tree as Record<string, unknown>)) {
      coerceCell(tree);
      return;
    }
    for (const rowKey of Object.keys(tree)) {
      const row = (tree as Record<string, unknown>)[rowKey];
      if (row == null || typeof row !== 'object') continue;
      if (looksLikeCell(row as Record<string, unknown>)) {
        coerceCell(row);
        continue;
      }
      for (const colKey of Object.keys(row as object)) {
        const cell = (row as Record<string, unknown>)[colKey];
        coerceCell(cell);
      }
    }
  }
}

const CELL_KEYS = new Set(['v', 'f', 't', 'p', 's', 'custom']);

function looksLikeCell(obj: Record<string, unknown>): boolean {
  // Leaf cells typically have only cell keys. Row maps have row-index
  // keys (strings of integers). Pure cell heuristic: ALL own keys must
  // be in CELL_KEYS. Empty objects are also treated as cells (Univer
  // uses `{}` to mean "clear").
  const keys = Object.keys(obj);
  if (keys.length === 0) return true;
  for (const k of keys) {
    if (!CELL_KEYS.has(k)) return false;
  }
  return true;
}

function coerceCell(cell: unknown): void {
  if (cell == null || typeof cell !== 'object') return;
  const c = cell as { v?: unknown; f?: unknown; t?: unknown };
  if (typeof c.f === 'string' && c.f.length > 0) return;
  if (typeof c.v !== 'string') return;
  const parsed = parseExcelStyleValue(c.v);
  if (parsed.kind !== 'number') return;
  c.v = parsed.value;
  // Clear any STRING / FORCE_STRING type marker so Univer infers
  // CellValueType.NUMBER and the "stored as text" alert doesn't fire.
  // `t` is left undefined; Univer derives it from `v` typeof at render.
  if ('t' in c) delete (c as { t?: unknown }).t;
}
