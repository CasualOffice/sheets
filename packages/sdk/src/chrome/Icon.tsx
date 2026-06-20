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
