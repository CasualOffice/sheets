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
 * Typed wrappers over the small subset of Univer's runtime API we
 * actually use. Stream B1 of the production-readiness pipeline.
 *
 * Why this exists:
 *
 * Univer's `@univerjs/core/facade` exports `FUniver`, `FWorkbook`,
 * `FWorksheet`, `FRange` — but the typed methods on those facade
 * classes don't cover every call we make from the shell + collab
 * layers. So callers reach for `as any` to call `getSheetId()`,
 * `isSheetHidden()`, `getMaxRows()`, `setActiveSheet()`, etc., and
 * the audit found ~150 `as any` sites concentrated at that boundary.
 *
 * Every `as any` is a lost type check. A facade refactor upstream
 * would invalidate every call silently; a typo in a method name
 * compiles fine and fails at runtime. We've been bitten by this
 * (e.g. the `_injector` access pattern that broke on a Univer
 * minor bump and surfaced as "bridge silently no-ops").
 *
 * This module narrows the boundary: ONE place where `as any` lives,
 * gated by typed wrappers that callers consume. When Univer's
 * facade changes, the type errors land here in one file instead
 * of scattered across the codebase.
 *
 * Scope: ONLY the methods we actually call. This is not a
 * comprehensive Univer typings package — it's a "what does Casual
 * Sheets need" surface. Add wrappers as new call sites appear;
 * don't pre-add coverage for hypothetical use.
 */
import type { FUniver } from '@univerjs/core/facade';
import type { FWorkbook, FWorksheet, FRange } from '@univerjs/sheets/facade';
import type { IWorkbookData, IRange } from '@univerjs/core';

/* ── Sheet identity + metadata ──────────────────────────────────── */

/**
 * Sheet id — the deterministic key (e.g. "sheet-1") used to address
 * the sheet across the collab op-log and the workbook snapshot. NOT
 * the human-facing name (use `sheet.getName()` for that).
 *
 * The facade types declare `getSheetId()` returns string, but it can
 * actually return `undefined` if the facade is wrapping a partially-
 * initialised worksheet — we narrow to `string | null` to force
 * callers to handle the missing case.
 */
export function sheetId(sheet: FWorksheet): string | null {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const id = (sheet as any).getSheetId?.();
  return typeof id === 'string' ? id : null;
}

/**
 * Whether this sheet is hidden from the tab strip. Hidden sheets are
 * still addressable by id and their formulas still resolve. Returns
 * false if the method isn't available (defensive — older Univer
 * builds shipped without this on the facade).
 */
export function isHidden(sheet: FWorksheet): boolean {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return (sheet as any).isSheetHidden?.() === true;
}

/**
 * Workbook capacity — Univer pre-allocates a fixed row/col count per
 * sheet (defaults: 1024 rows × 128 cols, configurable per workbook).
 * Used by the "select entire row/column" extenders so they don't
 * walk past the allocated bounds.
 */
export function maxRows(sheet: FWorksheet, fallback = 1024): number {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const n = (sheet as any).getMaxRows?.();
  return typeof n === 'number' && Number.isFinite(n) ? n : fallback;
}

export function maxColumns(sheet: FWorksheet, fallback = 128): number {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const n = (sheet as any).getMaxColumns?.();
  return typeof n === 'number' && Number.isFinite(n) ? n : fallback;
}

/* ── Range construction + activation ────────────────────────────── */

/**
 * Get a range by row/col rectangle. The facade types accept either
 * `(row, col)` for a single cell OR `IRange` for a rectangle, but
 * not both via the same overload — wrap so callers don't have to
 * juggle which signature exists in the current Univer version.
 */
export function rangeAt(sheet: FWorksheet, row: number, col: number): FRange | null {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (sheet as any).getRange?.(row, col) ?? null;
  } catch {
    return null;
  }
}

export function rangeBox(sheet: FWorksheet, box: IRange): FRange | null {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (sheet as any).getRange?.(box) ?? null;
  } catch {
    return null;
  }
}

/**
 * Resolve an A1 reference ("B2", "Sheet2!C5:D10") to a range on the
 * given sheet. Returns null when the reference doesn't parse OR
 * targets a different sheet that doesn't exist. The facade's
 * `getRange(string)` overload throws on bad input; we catch +
 * swallow so callers can prompt the user instead of crashing.
 */
export function rangeFromA1(sheet: FWorksheet, a1: string): FRange | null {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (sheet as any).getRange?.(a1) ?? null;
  } catch {
    return null;
  }
}

/**
 * Make this range the active selection on its sheet. Bridges the
 * gap between the facade's `range.activate()` (which exists on FRange
 * but isn't always exposed in older typings) and the imperative
 * "select this cell" UX the shell needs.
 */
export function activateRange(range: FRange): void {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (range as any).activate?.();
}

/* ── Active selection / data extents ────────────────────────────── */

/**
 * Get the range currently used as a data extent on the sheet — i.e.
 * the bounding box of cells that contain data. Used by Ctrl+End,
 * Flash Fill, and the pivot drill-down to know where the work is.
 * Falls back to the active range when the data range isn't available
 * (e.g. an empty sheet).
 */
export function dataRangeOrActive(sheet: FWorksheet): FRange | null {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const s = sheet as any;
  return s.getDataRange?.() ?? s.getActiveRange?.() ?? null;
}

/* ── Workbook-level operations ──────────────────────────────────── */

/**
 * Set the active sheet — wraps the facade's setter so the optional-
 * chain pattern at the call site (`(wb as any).setActiveSheet?.(s)`)
 * disappears.
 */
export function setActiveSheet(wb: FWorkbook, sheet: FWorksheet): void {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (wb as any).setActiveSheet?.(sheet);
}

/**
 * Find a sheet by its id. Returns null when no match. Wraps the
 * `getSheets().find()` pattern with `as any` on each iteration.
 */
export function findSheetById(wb: FWorkbook, id: string): FWorksheet | null {
  const sheets = wb.getSheets() as FWorksheet[];
  for (const s of sheets) {
    if (sheetId(s) === id) return s;
  }
  return null;
}

/**
 * Serialize the workbook to its IWorkbookData snapshot — the JSON-
 * able form used for upload to /snapshot, Y.Doc compaction, and
 * version history. The facade's `save()` returns `unknown` (Univer's
 * own typing); we narrow to the actual shape we expect.
 */
export function saveWorkbook(wb: FWorkbook): IWorkbookData | null {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const data = (wb as any).save?.();
    return data ?? null;
  } catch {
    return null;
  }
}

/* ── FUniver-level (top of the chain) ───────────────────────────── */

/**
 * Active sheet on the active workbook — the most common "get me
 * something to operate on" entry point. Returns null when no
 * workbook is mounted (uncommon but possible during teardown).
 */
export function activeSheet(api: FUniver): FWorksheet | null {
  return api.getActiveWorkbook()?.getActiveSheet() ?? null;
}

/**
 * Active range on the active sheet. Convenience for `activeSheet(api)?.getActiveRange()`.
 */
export function activeRange(api: FUniver): FRange | null {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return (activeSheet(api) as any)?.getActiveRange?.() ?? null;
}

/* ── Internal injector access (advanced) ────────────────────────── */

/**
 * Untyped Univer DI injector. Bridge code reaches into this to get
 * services like ICommandService that the public facade doesn't
 * expose. Every consumer that uses this should:
 *   1. Add a TODO referencing this fn so we know to revisit when
 *      Univer adds a facade-level alternative.
 *   2. Narrow the result with `instanceof` checks where possible.
 */
export interface UniverInjector {
  get<T = unknown>(token: unknown): T;
}

export function injector(api: FUniver): UniverInjector | null {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const inj = (api as any)._injector;
  if (!inj || typeof inj.get !== 'function') return null;
  return inj as UniverInjector;
}

/* ── Vite env helper ───────────────────────────────────────────── */

/**
 * Read a Vite env var with a default. `import.meta.env` types as
 * `ImportMetaEnv` which only knows about Vite's built-ins — our
 * VITE_* vars require an `as any` to access. Centralise here so we
 * have ONE place to either augment ImportMetaEnv OR keep the cast.
 */
export function viteEnv(key: string, fallback?: string): string | undefined {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const v = (import.meta.env as any)[key];
  if (v === undefined || v === null || v === '') return fallback;
  return String(v);
}

export function viteEnvNumber(key: string, fallback: number): number {
  const raw = viteEnv(key);
  if (raw === undefined) return fallback;
  const n = Number(raw);
  return Number.isFinite(n) ? n : fallback;
}

/* ── Window globals ─────────────────────────────────────────────── */

/**
 * Read a string-valued window global (e.g. `__COLLAB_WS_URL__` set
 * by an integration host before our bundle loads). Replaces the
 * `(window as any).__X__` pattern with a typed accessor.
 *
 * Only accepts string values; returns undefined for missing /
 * non-string globals. Callers that need other types should add a
 * specific helper rather than expanding this one.
 */
export function windowStringGlobal(key: string): string | undefined {
  if (typeof window === 'undefined') return undefined;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const v = (window as any)[key];
  return typeof v === 'string' && v.length > 0 ? v : undefined;
}
