import type { FUniver } from '@univerjs/core/facade';

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
