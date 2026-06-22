import { BorderStyleTypes, BorderType } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import { ensurePluginByName } from '@casualoffice/sheets/univer';

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

export function setBold(api: FUniver, bold: boolean) {
  activeRange(api)?.setFontWeight(bold ? 'bold' : 'normal');
}

export function toggleItalic(api: FUniver, currentlyItalic: boolean) {
  activeRange(api)?.setFontStyle(currentlyItalic ? 'normal' : 'italic');
}

export function setItalic(api: FUniver, italic: boolean) {
  activeRange(api)?.setFontStyle(italic ? 'italic' : 'normal');
}

export function toggleUnderline(api: FUniver, currentlyUnderline: boolean) {
  activeRange(api)?.setFontLine(currentlyUnderline ? 'none' : 'underline');
}

export function setUnderline(api: FUniver, underline: boolean) {
  activeRange(api)?.setFontLine(underline ? 'underline' : 'none');
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

/**
 * Text rotation in degrees, -90..90 (0 = horizontal). Maps to Univer's cell
 * `tr.a` angle via the sheets `FRange.setTextRotation` facade extension.
 */
export function setTextRotation(api: FUniver, degrees: number) {
  const range = activeRange(api) as unknown as {
    setTextRotation?: (d: number) => unknown;
  } | null;
  range?.setTextRotation?.(degrees);
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

export function setWrap(api: FUniver, wrapped: boolean) {
  activeRange(api)?.setWrap(wrapped);
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

export function pasteFormattingOnly(api: FUniver) {
  api.executeCommand('sheet.command.paste-format');
}

/**
 * Excel's Paste Special. Each mode maps to one of Univer's predefined
 * paste hooks — the base `univer.command.paste` command accepts a
 * `value` param naming the hook (see
 * vendor/univer/packages/sheets-ui/src/commands/commands/clipboard.command.ts
 * and the PREDEFINED_HOOK_NAME_PASTE table in clipboard.service.ts).
 *
 * Every mode resolves to a normal `set-range-values` / `set-style`
 * mutation, so co-edit propagation and xlsx round-trip are automatic —
 * no special handling needed.
 */
export type PasteSpecialMode =
  | 'all'
  | 'formulas'
  | 'values'
  | 'formats'
  | 'col-widths'
  | 'no-borders';

const PASTE_HOOK_VALUE: Record<PasteSpecialMode, string> = {
  all: 'default-paste',
  formulas: 'special-paste-formula',
  values: 'special-paste-value',
  formats: 'special-paste-format',
  'col-widths': 'special-paste-col-width',
  'no-borders': 'special-paste-besides-border',
};

export function pasteSpecial(api: FUniver, mode: PasteSpecialMode) {
  api.executeCommand('univer.command.paste', { value: PASTE_HOOK_VALUE[mode] });
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
 * Apply borders to the active selection. Default matches Excel's
 * default thin-border colour (black) so applied borders are clearly
 * distinguishable from Univer's `#dee0e3` gridlines — previously
 * `#666666`, which on a white grid at 100% zoom was indistinguishable
 * from gridlines after subpixel antialiasing.
 */
export type BorderChoice = 'all' | 'outside' | 'top' | 'bottom' | 'left' | 'right' | 'none';

/**
 * Three line weights matching Excel's default border dropdown:
 * thin / medium / thick. We don't expose Univer's dashed/dotted/double
 * variants here — those live in the Format Cells dialog where the
 * user has more room to pick. The weight is sticky for the session,
 * same as the colour.
 */
export type BorderWeight = 'thin' | 'medium' | 'thick';

export const DEFAULT_BORDER_COLOR = '#000000';
export const DEFAULT_BORDER_WEIGHT: BorderWeight = 'thin';

const WEIGHT_TO_STYLE: Record<BorderWeight, BorderStyleTypes> = {
  thin: BorderStyleTypes.THIN,
  medium: BorderStyleTypes.MEDIUM,
  thick: BorderStyleTypes.THICK,
};

export function setBorders(
  api: FUniver,
  choice: BorderChoice,
  color: string = DEFAULT_BORDER_COLOR,
  weight: BorderWeight = DEFAULT_BORDER_WEIGHT,
) {
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
  // "No border" ignores both weight and colour — the cell loses every
  // side. Otherwise map the picked weight to a Univer enum.
  const style = choice === 'none' ? BorderStyleTypes.NONE : WEIGHT_TO_STYLE[weight];
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
