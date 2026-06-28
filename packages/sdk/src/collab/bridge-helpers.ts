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
 * Pure helpers extracted from bridge.ts so they're testable without
 * pulling Univer / Yjs / Vite-only globals into a Node test runner.
 *
 * Anything stateless and Univer-free that bridge.ts needs should live
 * here. Stateful bridge logic (Yjs observers, command service hooks,
 * compaction) stays in bridge.ts itself.
 */

/**
 * Substitute `localUnitId` for every `unitId` string property anywhere
 * in `value` — including nested objects and arrays — and return a
 * structurally cloned copy. Returns the input by reference when no
 * change is needed, so callers can identity-compare before re-encoding.
 *
 * Walks plain objects and arrays only. Class instances are left alone
 * because Univer mutation params are required to be JSON-friendly.
 */
/**
 * Drawing mutations (sheets-drawing + drawing plugins) carry an
 * `op` field that's a json1 patch — a positional array whose first
 * element is the unitId. `deepRewriteUnitId` only rewrites the
 * `unitId` KEY in objects; the json1 path is a bare array of strings
 * so the unitId at position 0 stays as the OWNER's id and the apply
 * fails on the joiner with no useful message ("Error" at
 * json1.type.apply).
 *
 * This walks the op shape — single JSONOp (array of mixed strings +
 * numbers + a final mutation component object) OR JSONOpList (array
 * of JSONOps) — and substitutes the leading unitId where it matches.
 * Returns a fresh structure on change so peers can identity-compare.
 *
 * Scope deliberately narrow: only checks element [0] of each op
 * (where the unitId always lives in our path schema). A deeper
 * path-rewrite would invent semantics the Univer source doesn't
 * document.
 */
export function rewriteJson1OpPathUnitId(
  op: unknown,
  oldUnitId: string,
  newUnitId: string,
): unknown {
  if (oldUnitId === newUnitId) return op;
  if (!Array.isArray(op)) return op;
  // Distinguish single JSONOp (path...component) from JSONOpList
  // (array of JSONOps). A JSONOp's elements are strings, numbers, or
  // a single trailing component object. A JSONOpList's elements are
  // themselves arrays.
  const looksLikeOpList = op.length > 0 && Array.isArray(op[0]);
  if (looksLikeOpList) {
    let changed = false;
    const next = op.map((entry) => {
      const r = rewriteJson1OpPathUnitId(entry, oldUnitId, newUnitId);
      if (r !== entry) changed = true;
      return r;
    });
    return changed ? next : op;
  }
  // Single JSONOp — substitute element [0] if it matches.
  if (op[0] === oldUnitId) {
    const next = [...op];
    next[0] = newUnitId;
    return next;
  }
  return op;
}

export function deepRewriteUnitId(value: unknown, localUnitId: string): unknown {
  if (Array.isArray(value)) {
    let changed = false;
    const next = value.map((item) => {
      const r = deepRewriteUnitId(item, localUnitId);
      if (r !== item) changed = true;
      return r;
    });
    return changed ? next : value;
  }
  if (value && typeof value === 'object' && Object.getPrototypeOf(value) === Object.prototype) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const obj = value as Record<string, any>;
    let changed = false;
    const out: Record<string, unknown> = {};
    for (const key of Object.keys(obj)) {
      const v = obj[key];
      if (key === 'unitId' && typeof v === 'string' && v !== localUnitId) {
        out[key] = localUnitId;
        changed = true;
      } else {
        const r = deepRewriteUnitId(v, localUnitId);
        if (r !== v) changed = true;
        out[key] = r;
      }
    }
    return changed ? out : value;
  }
  return value;
}
