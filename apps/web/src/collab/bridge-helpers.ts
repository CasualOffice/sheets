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
