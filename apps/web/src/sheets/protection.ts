/**
 * Sheet + range protection (T4.4). Thin wrappers over Univer's
 * worksheet-permission facade (`FWorksheet.getWorksheetPermission()`), which
 * owns the model, the edit veto, the locked rendering, and xlsx persistence:
 *   - range:  protect a selection (rest of the sheet stays editable),
 *   - sheet:  protect the whole active sheet (sibling sheets stay editable).
 *
 * Collab model (chosen over Excel's block-everyone): the protector keeps
 * editing; OTHER editors are blocked. So `isProtected()` is the truth signal,
 * not the local user's `canEditCell`. Distinct from the workbook "Make
 * read-only" toggle, which flips the whole workbook via `applyReadOnly`.
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

/** Whether the active sheet currently has worksheet protection on. */
export function isActiveSheetProtected(api: FUniver): boolean {
  const perm = worksheetPermission(api);
  try {
    return perm?.isProtected?.() === true;
  } catch {
    return false;
  }
}

/**
 * Protect the whole active sheet (collab model: other editors can't change it;
 * the protector still can). Per-sheet — sibling sheets stay editable. Returns
 * true on success.
 */
export async function protectActiveSheet(api: FUniver): Promise<boolean> {
  const perm = worksheetPermission(api);
  if (!perm?.protect) return false;
  try {
    await perm.protect();
    return true;
  } catch {
    return false;
  }
}

/** Lift worksheet protection from the active sheet. Returns true on success. */
export async function unprotectActiveSheet(api: FUniver): Promise<boolean> {
  const perm = worksheetPermission(api);
  if (!perm?.unprotect) return false;
  try {
    await perm.unprotect();
    return true;
  } catch {
    return false;
  }
}
