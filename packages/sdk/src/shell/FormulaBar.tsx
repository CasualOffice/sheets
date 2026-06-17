import type { CSSProperties } from 'react';
import { Icon } from '@schnsrw/design-system';

export interface FormulaBarProps {
  /** Cell address, e.g. "A1" or a named range. */
  address?: string;
  /** Cell value (formula source or literal). */
  value?: string;
  /** Name Box click — opens Name Manager or namable list. */
  onOpenNameBox?: () => void;
  /** Drives `value` editing. */
  onValueChange?: (next: string) => void;
  onValueCommit?: (next: string) => void;
  onValueCancel?: () => void;
  readOnly?: boolean;
  style?: CSSProperties;
}

export function FormulaBar({
  address = 'A1',
  value = '',
  onOpenNameBox,
  onValueChange,
  onValueCommit,
  onValueCancel,
  readOnly = false,
  style,
}: FormulaBarProps) {
  return (
    <div
      style={{
        height: 34,
        flex: '0 0 34px',
        display: 'flex',
        alignItems: 'center',
        gap: 6,
        background: 'var(--color-surface)',
        borderBottom: '1px solid var(--color-divider)',
        padding: '0 8px',
        ...style,
      }}
    >
      <button
        type="button"
        onClick={onOpenNameBox}
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: 6,
          width: 104,
          height: 24,
          padding: '0 8px',
          borderRadius: 'var(--radius-md)',
          border: '1px solid var(--color-border)',
          background: 'var(--color-surface-alt)',
          fontFamily: 'var(--font-mono)',
          fontSize: 'var(--text-sm)',
          color: 'var(--color-text)',
          cursor: onOpenNameBox ? 'pointer' : 'default',
        }}
      >
        <span>{address}</span>
        <Icon
          name="arrow_drop_down"
          size={16}
          style={{ color: 'var(--color-text-muted)' } as CSSProperties}
        />
      </button>
      <span style={{ width: 1, height: 18, background: 'var(--color-divider)' }} />
      <span
        style={{
          fontFamily: 'var(--font-mono)',
          fontStyle: 'italic',
          fontSize: 'var(--text-base)',
          color: 'var(--color-text-secondary)',
          padding: '0 6px',
        }}
      >
        fx
      </span>
      <input
        type="text"
        value={value}
        readOnly={readOnly}
        onChange={(e) => onValueChange?.(e.currentTarget.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            onValueCommit?.(e.currentTarget.value);
          } else if (e.key === 'Escape') {
            e.preventDefault();
            onValueCancel?.();
          }
        }}
        style={{
          flex: 1,
          height: 24,
          padding: '0 10px',
          border: 0,
          outline: 0,
          background: 'transparent',
          fontFamily: 'var(--font-mono)',
          fontSize: 'var(--text-sm)',
          color: 'var(--color-text)',
        }}
      />
    </div>
  );
}
