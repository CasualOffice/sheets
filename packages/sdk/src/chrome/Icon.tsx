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

/**
 * Icon — a Material Symbols Outlined glyph for the chrome.
 *
 * Names are Material Symbols ids (https://fonts.google.com/icons), e.g.
 * `format_bold`. Requires the Material Symbols font (loaded via
 * `ensureChromeFonts`); if it isn't present the ligature text shows instead,
 * which is harmless. Icon-only buttons must carry their own `aria-label`.
 */

import type { CSSProperties } from 'react';

export interface IconProps {
  name: string;
  size?: number;
  style?: CSSProperties;
}

export function Icon({ name, size = 20, style }: IconProps) {
  return (
    <span
      className="material-symbols-outlined"
      aria-hidden
      style={{
        fontFamily: "'Material Symbols Outlined'",
        fontWeight: 'normal',
        fontStyle: 'normal',
        fontSize: size,
        lineHeight: 1,
        letterSpacing: 'normal',
        textTransform: 'none',
        display: 'inline-block',
        whiteSpace: 'nowrap',
        wordWrap: 'normal',
        direction: 'ltr',
        // Hint the variable-font axes so weight/optical-size are consistent.
        fontVariationSettings: "'opsz' 20, 'wght' 400, 'FILL' 0, 'GRAD' 0",
        ...style,
      }}
    >
      {name}
    </span>
  );
}
