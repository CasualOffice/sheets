import type { FUniver } from '@univerjs/core/facade';

/** Sheet-tab imperative actions. */

export function switchToSheet(api: FUniver, sheetId: string) {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  const target = wb.getSheets().find((s) => s.getSheetId() === sheetId);
  if (!target) return;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (wb as any).setActiveSheet(target);
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
  sheet.setName(name.slice(0, 31));
  return true;
}

export function deleteSheetById(api: FUniver, sheetId: string): boolean {
  const wb = api.getActiveWorkbook();
  if (!wb) return false;
  if (wb.getSheets().length <= 1) return false; // can't delete the last sheet
  return wb.deleteSheet(sheetId);
}

export function addSheet(api: FUniver) {
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
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (target as any).hideSheet?.();
  return true;
}

/** Show a previously-hidden sheet. */
export function showSheet(api: FUniver, sheetId: string): boolean {
  const wb = api.getActiveWorkbook();
  if (!wb) return false;
  const target = wb.getSheets().find((s) => s.getSheetId() === sheetId);
  if (!target) return false;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (target as any).showSheet?.();
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
