import { BorderStyleTypes, BorderType } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import { ensurePluginByName } from '../univer/lazy-plugins';

/**
 * Imperative command dispatchers for Home-tab buttons.
 * Kept separate from the Ribbon component so the same actions can be reused
 * by keyboard shortcuts (Phase 1.4) without duplicating logic.
 */

function activeRange(api: FUniver) {
  const wb = api.getActiveWorkbook();
  const sheet = wb?.getActiveSheet();
  return sheet?.getActiveRange() ?? null;
}

export function toggleBold(api: FUniver, currentlyBold: boolean) {
  activeRange(api)?.setFontWeight(currentlyBold ? 'normal' : 'bold');
}

export function toggleItalic(api: FUniver, currentlyItalic: boolean) {
  activeRange(api)?.setFontStyle(currentlyItalic ? 'normal' : 'italic');
}

export function toggleUnderline(api: FUniver, currentlyUnderline: boolean) {
  activeRange(api)?.setFontLine(currentlyUnderline ? 'none' : 'underline');
}

/**
 * Univer's Facade API uses 'normal' to mean right-aligned (see
 * `vendor/univer/packages/sheets/src/facade/utils.ts` — `FHorizontalAlignment`).
 * We keep the user-facing alignment names ('left' | 'center' | 'right') and
 * translate at the boundary.
 */
export function setAlignment(api: FUniver, alignment: 'left' | 'center' | 'right') {
  const facadeValue = alignment === 'right' ? 'normal' : alignment;
  activeRange(api)?.setHorizontalAlignment(facadeValue);
}

export function setNumberFormat(api: FUniver, pattern: string) {
  // setNumberFormat lives in the sheets-numfmt facade extension.
  // It augments FRange at runtime via FUniver.extend(), so a runtime cast is
  // the cleanest way to use it without re-declaring the type surface here.
  const range = activeRange(api) as unknown as
    | { setNumberFormat?: (p: string) => unknown }
    | null;
  range?.setNumberFormat?.(pattern);
}

export const NUMBER_FORMATS = {
  currency: '"$"#,##0.00',
  percent: '0.00%',
} as const;

/**
 * Vertical alignment — `setVerticalAlignment` accepts 'top' | 'middle' | 'bottom'.
 */
export function setVerticalAlignment(api: FUniver, alignment: 'top' | 'middle' | 'bottom') {
  activeRange(api)?.setVerticalAlignment(alignment);
}

export function setFontFamily(api: FUniver, family: string) {
  activeRange(api)?.setFontFamily(family || null);
}

export function setFontSize(api: FUniver, size: number) {
  if (!Number.isFinite(size) || size <= 0) return;
  activeRange(api)?.setFontSize(size);
}

/**
 * Bump the active cell's font size by `delta` (commonly +1 / -1 from
 * the A↑ / A↓ buttons). Reads the current size off the active cell's
 * style; falls back to 11 (Excel default) when no explicit size set.
 * Clamped to the same [6, 72] window the font-size dropdown offers.
 */
export function adjustFontSize(api: FUniver, delta: number) {
  const wb = api.getActiveWorkbook();
  const sheet = wb?.getActiveSheet();
  const range = sheet?.getActiveRange();
  if (!wb || !sheet || !range) return;
  const cell = sheet.getRange(range.getRow(), range.getColumn());
  const data = cell.getCellData();
  const style =
    typeof data?.s === 'string'
      ? (wb.getWorkbook().getStyles().get(data.s) ?? null)
      : (data?.s ?? null);
  const current = typeof style?.fs === 'number' && style.fs > 0 ? style.fs : 11;
  const next = Math.max(6, Math.min(72, current + delta));
  if (next === current) return;
  setFontSize(api, next);
}

export function setFontColor(api: FUniver, color: string) {
  activeRange(api)?.setFontColor(color || null);
}

export function setFillColor(api: FUniver, color: string) {
  activeRange(api)?.setBackground(color);
}

export function toggleWrap(api: FUniver, currentlyWrapped: boolean) {
  activeRange(api)?.setWrap(!currentlyWrapped);
}

/**
 * Strikethrough goes through the Facade's `setFontLine` since the dedicated
 * setter is private — passing 'line-through' enables it.
 */
export function toggleStrikethrough(api: FUniver, currentlyStrike: boolean) {
  // 'line-through' / 'none' for strike; the same facade method also handles
  // underline ('underline' / 'none'). Existing underline state is preserved
  // by the underlying SetStyleCommand because we only mutate the 'st' style key.
  activeRange(api)?.setFontLine(currentlyStrike ? 'none' : 'line-through');
}

/* ── Undo / Redo / Clipboard ────────────────────────────────────────────── */

export function undo(api: FUniver) {
  api.executeCommand('univer.command.undo');
}

export function redo(api: FUniver) {
  api.executeCommand('univer.command.redo');
}

export function copy(api: FUniver) {
  api.executeCommand('univer.command.copy');
}

export function cut(api: FUniver) {
  api.executeCommand('univer.command.cut');
}

export function paste(api: FUniver) {
  api.executeCommand('univer.command.paste');
}

/* ── Number format helpers ──────────────────────────────────────────────── */

export const NUMBER_FORMAT_PATTERNS = {
  general: '',
  number: '#,##0.00',
  integer: '#,##0',
  currency: '"$"#,##0.00',
  accounting: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  percent: '0.00%',
  date: 'yyyy-mm-dd',
  time: 'hh:mm:ss',
  scientific: '0.00E+00',
  text: '@',
} as const;

export type NumberFormatKey = keyof typeof NUMBER_FORMAT_PATTERNS;

export function setNumberFormatByKey(api: FUniver, key: NumberFormatKey) {
  setNumberFormat(api, NUMBER_FORMAT_PATTERNS[key]);
}

export function increaseDecimal(api: FUniver) {
  api.executeCommand('sheet.command.numfmt.add.decimal.command');
}

export function decreaseDecimal(api: FUniver) {
  api.executeCommand('sheet.command.numfmt.subtract.decimal.command');
}

/* ── Format Painter ─────────────────────────────────────────────────────── */

/** One-shot painter: capture current style, next selection click applies it. */
export function startFormatPainter(api: FUniver) {
  api.executeCommand('sheet.command.set-once-format-painter');
}

/* ── Find & Replace ─────────────────────────────────────────────────────── */

export async function openFindReplace(api: FUniver) {
  // find-replace is lazy-loaded. Without the await, hitting Ctrl+F on
  // a fresh page (before the idle-load fires) dispatches the operation
  // to a not-yet-registered handler and silently no-ops. ensurePluginByName
  // is idempotent: it resolves immediately once the plugin is in.
  await ensurePluginByName('findReplace');
  // Operation id from @univerjs/find-replace
  // (find-replace.operation.ts:23 in the vendored source).
  api.executeCommand('ui.operation.open-find-dialog');
}

/**
 * Apply borders to the active selection. Color defaults to the standard
 * neutral grid color so borders blend with Excel-style sheets; pass an
 * explicit hex from the toolbar's color picker to override.
 */
export type BorderChoice = 'all' | 'outside' | 'top' | 'bottom' | 'left' | 'right' | 'none';

export const DEFAULT_BORDER_COLOR = '#666666';

export function setBorders(api: FUniver, choice: BorderChoice, color: string = DEFAULT_BORDER_COLOR) {
  const range = activeRange(api);
  if (!range) return;
  const type =
    choice === 'all'
      ? BorderType.ALL
      : choice === 'outside'
        ? BorderType.OUTSIDE
        : choice === 'top'
          ? BorderType.TOP
          : choice === 'bottom'
            ? BorderType.BOTTOM
            : choice === 'left'
              ? BorderType.LEFT
              : choice === 'right'
                ? BorderType.RIGHT
                : BorderType.NONE;
  const style = choice === 'none' ? BorderStyleTypes.NONE : BorderStyleTypes.THIN;
  range.setBorder(type, style, color);
}

/**
 * Toggle merge on the current selection. If the selection is a single
 * already-merged range, unmerge it; otherwise merge and center.
 */
export function toggleMerge(api: FUniver, currentlyMerged: boolean) {
  const range = activeRange(api);
  if (!range) return;
  if (currentlyMerged) {
    range.breakApart();
  } else {
    // Skip the 1x1 no-op — merging a single cell is meaningless.
    if (range.getWidth() * range.getHeight() <= 1) return;
    range.merge();
    range.setHorizontalAlignment('center');
  }
}
