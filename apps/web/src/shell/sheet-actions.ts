import type { FUniver } from '@univerjs/core/facade';

/**
 * Sheet-tab imperative actions.
 *
 * Every mutation that needs to propagate to co-edit peers MUST dispatch
 * through the command service (`api.executeCommand`), not via the
 * facade's direct methods (`sheet.setName`, `target.hideSheet`, …).
 * The collab bridge captures mutations via
 * `ICommandService.onMutationExecutedForCollab`; facade methods that
 * bypass the bus don't fire that hook, so renames/hides/activates
 * would silently stay local-only. See docs/COLLAB-FIXES.md issues
 * 21 + 22.
 */

export function switchToSheet(api: FUniver, sheetId: string) {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  // Goes through the command bus → SetWorksheetActiveOperation mutation
  // is captured by the bridge. The facade's `setActiveSheet` does NOT
  // emit through the command bus and would skip remote-sync.
  void api.executeCommand('sheet.command.set-worksheet-activate', {
    unitId: wb.getId(),
    subUnitId: sheetId,
  });
}

export function renameSheet(api: FUniver, sheetId: string, name: string): boolean {
  if (!name.trim()) return false;
  const wb = api.getActiveWorkbook();
  if (!wb) return false;
  const sheet = wb.getSheets().find((s) => s.getSheetId() === sheetId);
  if (!sheet) return false;
  // Sheet names in xlsx must be unique and ≤ 31 chars.
  const exists = wb
    .getSheets()
    .some((s) => s.getSheetId() !== sheetId && s.getSheetName() === name);
  if (exists) return false;
  void api.executeCommand('sheet.command.set-worksheet-name', {
    unitId: wb.getId(),
    subUnitId: sheetId,
    name: name.slice(0, 31),
  });
  return true;
}

export function deleteSheetById(api: FUniver, sheetId: string): boolean {
  const wb = api.getActiveWorkbook();
  if (!wb) return false;
  if (wb.getSheets().length <= 1) return false; // can't delete the last sheet
  // `wb.deleteSheet` dispatches `RemoveSheetCommand` internally — the
  // bridge picks up the resulting `sheet.mutation.remove-sheet`.
  return wb.deleteSheet(sheetId);
}

export function addSheet(api: FUniver) {
  // `insertSheet` dispatches `InsertSheetCommand` → mutation captured.
  api.getActiveWorkbook()?.insertSheet();
}

/** Hide the given sheet. Excel keeps the sheet in the workbook (and
 *  its formulas still resolve) — just not visible in the tab strip.
 *  No-ops if it's the only visible sheet; Excel disallows hiding the
 *  last one. */
export function hideSheet(api: FUniver, sheetId: string): boolean {
  const wb = api.getActiveWorkbook();
  if (!wb) return false;
  const sheets = wb.getSheets();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const visibleCount = sheets.filter((s) => (s as any).isSheetHidden?.() !== true).length;
  if (visibleCount <= 1) return false;
  const target = sheets.find((s) => s.getSheetId() === sheetId);
  if (!target) return false;
  void api.executeCommand('sheet.command.set-worksheet-hidden', {
    unitId: wb.getId(),
    subUnitId: sheetId,
    hidden: 1, // BooleanNumber.TRUE
  });
  return true;
}

/** Show a previously-hidden sheet. */
export function showSheet(api: FUniver, sheetId: string): boolean {
  const wb = api.getActiveWorkbook();
  if (!wb) return false;
  const target = wb.getSheets().find((s) => s.getSheetId() === sheetId);
  if (!target) return false;
  void api.executeCommand('sheet.command.set-worksheet-show', {
    unitId: wb.getId(),
    subUnitId: sheetId,
  });
  return true;
}

/**
 * Duplicate a sheet (Excel "Move or Copy → Create a copy"). Univer ships the
 * heavy lifting — we just dispatch the command with the source subUnitId.
 * The unit's active sheet ends up on the new copy.
 */
export function duplicateSheet(api: FUniver, sheetId: string): void {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  void api.executeCommand('sheet.command.copy-sheet', {
    unitId: wb.getId(),
    subUnitId: sheetId,
  });
}

/**
 * Move a sheet to a new index in the tab order.
 */
export function moveSheetTo(api: FUniver, sheetId: string, toIndex: number) {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  const sheets = wb.getSheets();
  const sheet = sheets.find((s) => s.getSheetId() === sheetId);
  if (!sheet) return;
  const clamped = Math.max(0, Math.min(toIndex, sheets.length - 1));
  wb.moveSheet(sheet, clamped);
}
