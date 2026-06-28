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

import type { CSSProperties } from 'react';

/**
 * Material Symbols Outlined icon — renders the exact same way as the
 * design-system's `Icon` (the UI-kit): the Material Symbols Outlined
 * variable font with axes `'FILL' x, 'wght' 400, 'GRAD' 0, 'opsz' 24`,
 * on the same size scale. This keeps ONE icon family across the app and
 * the DS components (Button / IconButton / Menu / …) which also render
 * the font — previously the app drew inline SVGs, a second family.
 *
 * The font is loaded by `@schnsrw/design-system/tokens.css` (imported at
 * the app entry). For icon-only buttons the parent carries the
 * `aria-label`; the glyph itself is decorative (`aria-hidden`).
 */
type IconSize = 'xs' | 'sm' | 'md' | 'lg' | 'xl';

type Props = {
  name: string;
  size?: IconSize | number;
  filled?: boolean;
  weight?: 100 | 200 | 300 | 400 | 500 | 600 | 700;
  className?: string;
  style?: CSSProperties;
};

// Matches the design-system Icon scale exactly.
const SIZE_PX: Record<IconSize, number> = {
  xs: 14,
  sm: 16,
  md: 18,
  lg: 20,
  xl: 24,
};

// A few names where we deliberately want a clearer glyph than the literal
// Material Symbols default — kept as font ligatures so they stay in-family.
const NAME_REMAP: Record<string, string> = {
  // Highlighter pen reads better than the paint-bucket for fill colour.
  format_color_fill: 'ink_highlighter',
};

export function Icon({ name, size = 'md', filled = false, weight = 400, className, style }: Props) {
  const px = typeof size === 'number' ? size : SIZE_PX[size];
  const ligature = NAME_REMAP[name] ?? name;
  const merged: CSSProperties = {
    fontSize: px,
    lineHeight: 1,
    fontVariationSettings: `'FILL' ${filled ? 1 : 0}, 'wght' ${weight}, 'GRAD' 0, 'opsz' 24`,
    ...style,
  };
  // Keep BOTH `material-symbols-outlined` (DS font class) and `icon` (the
  // app's existing selectors target `.icon` for colour/layout). Inline
  // fontSize/variation win over both, so sizing stays exact.
  const classes = `material-symbols-outlined icon${className ? ` ${className}` : ''}`;
  return (
    <span aria-hidden="true" className={classes} style={merged}>
      {ligature}
    </span>
  );
}
