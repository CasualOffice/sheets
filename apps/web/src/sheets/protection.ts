/**
 * Range protection (T4.4) — lock a selected range while the rest of the sheet
 * stays editable. Thin wrappers over Univer's worksheet-permission facade
 * (`FWorksheet.getWorksheetPermission()`), which owns the model, the edit veto,
 * and the locked-range rendering. We only drive it from the Review menu.
 *
 * This is finer-grained than the existing workbook "Protect (read-only)" toggle
 * (which flips the whole workbook via `applyReadOnly`): here only the chosen
 * cells refuse edits.
 */
import type { FUniver } from '@univerjs/core/facade';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyWs = any;

function activeWorksheet(api: FUniver): AnyWs | null {
  try {
    return api.getActiveWorkbook()?.getActiveSheet() ?? null;
  } catch {
    return null;
  }
}

function worksheetPermission(api: FUniver): AnyWs | null {
  const ws = activeWorksheet(api);
  return ws?.getWorksheetPermission?.() ?? null;
}

export type ProtectRangeResult =
  | { ok: true; a1: string }
  | { ok: false; reason: 'no-selection' | 'overlap' | 'unavailable' };

/**
 * Protect the active selection. Returns the A1 of the protected range on
 * success, or a reason the caller can surface as a toast. Overlap with an
 * existing rule is reported rather than thrown (the facade throws on overlap).
 */
export async function protectActiveRange(api: FUniver): Promise<ProtectRangeResult> {
  const ws = activeWorksheet(api);
  const perm = ws?.getWorksheetPermission?.();
  if (!ws || !perm?.protectRanges) return { ok: false, reason: 'unavailable' };
  const range = ws.getActiveRange?.();
  if (!range) return { ok: false, reason: 'no-selection' };
  try {
    await perm.protectRanges([{ ranges: [range] }]);
    return { ok: true, a1: range.getA1Notation?.() ?? '' };
  } catch (e) {
    if (e instanceof Error && /overlap/i.test(e.message)) return { ok: false, reason: 'overlap' };
    return { ok: false, reason: 'unavailable' };
  }
}

/** Number of range-protection rules on the active sheet. */
export async function countRangeProtections(api: FUniver): Promise<number> {
  const perm = worksheetPermission(api);
  if (!perm?.listRangeProtectionRules) return 0;
  try {
    const rules = await perm.listRangeProtectionRules();
    return Array.isArray(rules) ? rules.length : 0;
  } catch {
    return 0;
  }
}

/**
 * Remove every range-protection rule on the active sheet. Returns how many were
 * cleared (0 if none / unavailable).
 */
export async function clearRangeProtections(api: FUniver): Promise<number> {
  const perm = worksheetPermission(api);
  if (!perm?.listRangeProtectionRules || !perm?.unprotectRules) return 0;
  try {
    const rules = await perm.listRangeProtectionRules();
    const ids: string[] = (rules ?? [])
      .map((r: AnyWs) => r?.id)
      .filter((id: unknown): id is string => typeof id === 'string' && id.length > 0);
    if (ids.length === 0) return 0;
    await perm.unprotectRules(ids);
    return ids.length;
  } catch {
    return 0;
  }
}
