import type { CSSProperties } from 'react';

/**
 * Material Symbols Outlined icon. Names are from
 * https://fonts.google.com/icons (e.g. "format_bold").
 *
 * For icon-only buttons, ensure the parent `<button>` carries an `aria-label`
 * — the icon itself is decorative.
 */
type Props = {
  name: string;
  size?: 'sm' | 'md' | 'lg';
  filled?: boolean;
  className?: string;
  style?: CSSProperties;
};

export function Icon({ name, size = 'md', filled, className, style }: Props) {
  const sizeClass = size === 'sm' ? ' icon--sm' : size === 'lg' ? ' icon--lg' : '';
  const filledClass = filled ? ' icon--filled' : '';
  return (
    <span
      className={`icon${sizeClass}${filledClass}${className ? ` ${className}` : ''}`}
      style={style}
      aria-hidden="true"
    >
      {name}
    </span>
  );
}
