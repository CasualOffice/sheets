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

import type { IStyleData } from '@univerjs/core';
import { BorderStyleTypes } from '@univerjs/core';
import type * as ExcelJS from 'exceljs';

// ExcelJS border-style strings ↔ Univer BorderStyleTypes. Covers Excel's full
// set of line styles so dashed / double / thick / medium / hair / dotted survive
// the round-trip instead of all collapsing to a thin line. Anything unrecognised
// falls back to THIN so a border is never dropped entirely.
const EXCEL_BORDER_TO_UNIVER: Record<string, BorderStyleTypes> = {
  thin: BorderStyleTypes.THIN,
  hair: BorderStyleTypes.HAIR,
  dotted: BorderStyleTypes.DOTTED,
  dashed: BorderStyleTypes.DASHED,
  dashDot: BorderStyleTypes.DASH_DOT,
  dashDotDot: BorderStyleTypes.DASH_DOT_DOT,
  double: BorderStyleTypes.DOUBLE,
  medium: BorderStyleTypes.MEDIUM,
  mediumDashed: BorderStyleTypes.MEDIUM_DASHED,
  mediumDashDot: BorderStyleTypes.MEDIUM_DASH_DOT,
  mediumDashDotDot: BorderStyleTypes.MEDIUM_DASH_DOT_DOT,
  slantDashDot: BorderStyleTypes.SLANT_DASH_DOT,
  thick: BorderStyleTypes.THICK,
};
const UNIVER_BORDER_TO_EXCEL: Record<number, ExcelJS.BorderStyle> = Object.fromEntries(
  Object.entries(EXCEL_BORDER_TO_UNIVER).map(([excel, univer]) => [univer, excel]),
) as Record<number, ExcelJS.BorderStyle>;

/**
 * Mappings between ExcelJS and Univer style models.
 *
 * Univer's `IStyleData` uses two-letter compact keys (bl, it, ff, fs, cl, bg,
 * ht, vt, n, ul, bd) plus numeric enums for alignment. ExcelJS uses verbose
 * keys (font.bold, alignment.horizontal, numFmt, etc.). Round-trip preserves
 * the basics; advanced styling (gradients, patterns, complex borders) falls
 * back to defaults.
 */

const ARGB_RX = /^#?([0-9A-Fa-f]{6,8})$/;

function normalizeColor(argb: string | undefined): string | undefined {
  if (!argb) return undefined;
  const m = ARGB_RX.exec(argb);
  if (!m) return undefined;
  const hex = m[1];
  // ExcelJS uses ARGB; drop alpha if present.
  const rgb = hex.length === 8 ? hex.slice(2) : hex;
  return `#${rgb.toLowerCase()}`;
}

// Accept Univer's Nullable<string> (string | null | void) by treating any
// non-string input as undefined.
function toARGB(rgb: unknown): string | undefined {
  if (typeof rgb !== 'string' || !rgb) return undefined;
  const m = ARGB_RX.exec(rgb);
  if (!m) return undefined;
  const hex = m[1];
  const norm = hex.length === 8 ? hex : `FF${hex}`;
  return norm.toUpperCase();
}

// Univer HorizontalAlign: 0 = default, 1 = LEFT, 2 = CENTER, 3 = RIGHT
const H_ALIGN_FROM_EXCEL: Record<string, number> = {
  left: 1,
  center: 2,
  right: 3,
};
const H_ALIGN_TO_EXCEL: Record<number, string> = {
  1: 'left',
  2: 'center',
  3: 'right',
};

// Cell indentation. Excel's `indent` is a level (each unit ≈ 3 spaces, OOXML);
// Univer's `pd.l` is left padding in px (replacing the 2px default). One level
// ≈ ~10px is a faithful approximation; the same constants run both directions
// so the level round-trips exactly.
const DEFAULT_LEFT_PADDING_PX = 2;
const INDENT_STEP_PX = 10;

// Univer VerticalAlign: 0 = default, 1 = TOP, 2 = MIDDLE, 3 = BOTTOM
const V_ALIGN_FROM_EXCEL: Record<string, number> = {
  top: 1,
  middle: 2,
  bottom: 3,
};
const V_ALIGN_TO_EXCEL: Record<number, string> = {
  1: 'top',
  2: 'middle',
  3: 'bottom',
};

export function excelStyleToUniver(cell: ExcelJS.Cell): IStyleData | undefined {
  const s: IStyleData = {};

  if (cell.font) {
    if (cell.font.name) s.ff = cell.font.name;
    if (cell.font.size) s.fs = cell.font.size;
    if (cell.font.bold) s.bl = 1;
    if (cell.font.italic) s.it = 1;
    if (cell.font.underline) s.ul = { s: 1 };
    if (cell.font.strike) s.st = { s: 1 };
    const fc = normalizeColor((cell.font.color as { argb?: string } | undefined)?.argb);
    if (fc) s.cl = { rgb: fc };
  }

  if (cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern === 'solid') {
    const bg = normalizeColor((cell.fill.fgColor as { argb?: string } | undefined)?.argb);
    if (bg) s.bg = { rgb: bg };
  }

  if (cell.alignment) {
    const ha = cell.alignment.horizontal;
    if (ha && H_ALIGN_FROM_EXCEL[ha] !== undefined) s.ht = H_ALIGN_FROM_EXCEL[ha];
    const va = cell.alignment.vertical;
    if (va && V_ALIGN_FROM_EXCEL[va] !== undefined) s.vt = V_ALIGN_FROM_EXCEL[va];
    if (cell.alignment.wrapText) s.tb = 3; // WrapStrategy.WRAP
    // Text rotation. ExcelJS gives a signed degree (-90..90) or 'vertical'
    // (stacked). Univer's tr.a is the angle (same OOXML convention) and tr.v
    // flags the vertical/stacked mode.
    const rot = cell.alignment.textRotation as number | 'vertical' | undefined;
    if (rot === 'vertical') s.tr = { a: 0, v: 1 };
    else if (typeof rot === 'number' && rot !== 0) s.tr = { a: rot };
    // Indentation. Excel stores a level (each ≈ 3 spaces per OOXML); Univer
    // renders left padding in px (pd.l replaces the default). Map level → px so
    // indented/outline data keeps its hierarchy instead of flattening.
    const indent = cell.alignment.indent;
    if (typeof indent === 'number' && indent > 0) {
      s.pd = { l: DEFAULT_LEFT_PADDING_PX + indent * INDENT_STEP_PX };
    }
  }

  if (cell.numFmt) s.n = { pattern: cell.numFmt };

  // Borders — map each present side's line style (EXCEL_BORDER_TO_UNIVER) + color.
  if (cell.border) {
    const bd: NonNullable<IStyleData['bd']> = {};
    const sides: Array<['t' | 'b' | 'l' | 'r', 'top' | 'bottom' | 'left' | 'right']> = [
      ['t', 'top'],
      ['b', 'bottom'],
      ['l', 'left'],
      ['r', 'right'],
    ];
    for (const [k, key] of sides) {
      const side = (cell.border as Record<string, unknown>)[key] as
        | { style?: string; color?: { argb?: string } }
        | undefined;
      if (side?.style && side.style !== 'none') {
        bd[k] = {
          s: EXCEL_BORDER_TO_UNIVER[side.style] ?? BorderStyleTypes.THIN,
          cl: { rgb: normalizeColor(side.color?.argb) ?? '#666666' },
        };
      }
    }
    if (Object.keys(bd).length > 0) s.bd = bd;
  }

  return Object.keys(s).length > 0 ? s : undefined;
}

export function univerStyleToExcel(style: IStyleData): Partial<ExcelJS.Style> {
  const out: Partial<ExcelJS.Style> = {};

  const font: ExcelJS.Style['font'] = {};
  if (style.ff) font.name = style.ff;
  if (style.fs) font.size = style.fs;
  if (style.bl === 1) font.bold = true;
  if (style.it === 1) font.italic = true;
  if (style.ul?.s === 1) font.underline = true;
  if (style.st?.s === 1) font.strike = true;
  if (style.cl && typeof style.cl === 'object' && 'rgb' in style.cl) {
    const argb = toARGB(style.cl.rgb);
    if (argb) font.color = { argb };
  }
  if (Object.keys(font).length > 0) out.font = font;

  if (style.bg && typeof style.bg === 'object' && 'rgb' in style.bg) {
    const argb = toARGB(style.bg.rgb);
    if (argb) {
      out.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb },
      };
    }
  }

  const alignment: ExcelJS.Style['alignment'] = {};
  if (style.ht && H_ALIGN_TO_EXCEL[style.ht]) {
    alignment.horizontal = H_ALIGN_TO_EXCEL[style.ht] as ExcelJS.Style['alignment']['horizontal'];
  }
  if (style.vt && V_ALIGN_TO_EXCEL[style.vt]) {
    alignment.vertical = V_ALIGN_TO_EXCEL[style.vt] as ExcelJS.Style['alignment']['vertical'];
  }
  if (style.tb === 3) alignment.wrapText = true;
  if (style.tr) {
    if (style.tr.v === 1) alignment.textRotation = 'vertical';
    else if (typeof style.tr.a === 'number' && style.tr.a !== 0)
      alignment.textRotation = style.tr.a;
  }
  if (typeof style.pd?.l === 'number') {
    const level = Math.round((style.pd.l - DEFAULT_LEFT_PADDING_PX) / INDENT_STEP_PX);
    if (level > 0) alignment.indent = level;
  }
  if (Object.keys(alignment).length > 0) out.alignment = alignment;

  if (style.n?.pattern) out.numFmt = style.n.pattern;

  if (style.bd) {
    const border: ExcelJS.Style['border'] = {};
    const sides: Array<['t' | 'b' | 'l' | 'r', 'top' | 'bottom' | 'left' | 'right']> = [
      ['t', 'top'],
      ['b', 'bottom'],
      ['l', 'left'],
      ['r', 'right'],
    ];
    for (const [k, key] of sides) {
      const side = style.bd[k];
      if (!side) continue;
      const color = side.cl && 'rgb' in side.cl ? toARGB(side.cl.rgb) : undefined;
      (border as Record<string, unknown>)[key] = {
        style: UNIVER_BORDER_TO_EXCEL[side.s as number] ?? 'thin',
        ...(color ? { color: { argb: color } } : {}),
      };
    }
    if (Object.keys(border).length > 0) out.border = border;
  }

  return out;
}
