import { useState, type CSSProperties } from 'react';
import { Icon } from '@schnsrw/design-system';

export interface SheetDescriptor {
  name: string;
  /** Optional accent dot colour. Falls back to a rotating palette. */
  color?: string;
}

export interface SheetTabsProps {
  sheets: SheetDescriptor[];
  active: number;
  onSelect: (index: number) => void;
  onAdd?: () => void;
  style?: CSSProperties;
}

const PALETTE = ['var(--color-accent)', '#15803d', '#8a8886'];

export function SheetTabs({ sheets, active, onSelect, onAdd, style }: SheetTabsProps) {
  const [hoverAdd, setHoverAdd] = useState(false);
  return (
    <div
      style={{
        height: 38,
        flex: '0 0 38px',
        display: 'flex',
        alignItems: 'center',
        gap: 6,
        background: 'var(--color-surface)',
        borderTop: '1px solid var(--color-divider)',
        padding: '0 10px',
        ...style,
      }}
    >
      <button
        type="button"
        title="Add sheet"
        onClick={onAdd}
        onMouseEnter={() => setHoverAdd(true)}
        onMouseLeave={() => setHoverAdd(false)}
        style={{
          width: 28,
          height: 28,
          border: 0,
          borderRadius: 'var(--radius-md)',
          background: hoverAdd ? 'var(--color-hover)' : 'transparent',
          cursor: onAdd ? 'pointer' : 'not-allowed',
          color: 'var(--color-text-secondary)',
          display: 'inline-flex',
          alignItems: 'center',
          justifyContent: 'center',
        }}
      >
        <Icon name="add" size="lg" />
      </button>
      <span style={{ width: 1, height: 18, background: 'var(--color-divider)' }} />
      <div
        style={{ display: 'flex', alignItems: 'center', gap: 4, overflowX: 'auto', flex: 1 }}
      >
        {sheets.map((s, i) => (
          <Tab
            key={`${i}-${s.name}`}
            sheet={s}
            active={i === active}
            dotColor={s.color ?? PALETTE[i % PALETTE.length]}
            onClick={() => onSelect(i)}
          />
        ))}
      </div>
    </div>
  );
}

function Tab({
  sheet,
  active,
  dotColor,
  onClick,
}: {
  sheet: SheetDescriptor;
  active: boolean;
  dotColor: string;
  onClick: () => void;
}) {
  const [hover, setHover] = useState(false);
  return (
    <button
      type="button"
      onClick={onClick}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        gap: 7,
        height: 28,
        padding: '0 12px',
        border: active ? '1px solid var(--color-border)' : '1px solid transparent',
        borderRadius: 'var(--radius-pill)',
        background: active
          ? 'var(--color-surface-alt)'
          : hover
            ? 'var(--color-hover)'
            : 'transparent',
        boxShadow: active ? 'var(--shadow-1)' : 'none',
        font: 'inherit',
        fontFamily: 'var(--font-sans)',
        fontSize: 12.5,
        fontWeight: active ? 'var(--weight-semibold)' : 'var(--weight-medium)',
        color: active ? 'var(--color-text)' : 'var(--color-text-secondary)',
        cursor: 'pointer',
        transition:
          'background var(--motion-fast) var(--ease-out), color var(--motion-fast) var(--ease-out)',
      }}
    >
      <span
        style={{
          width: 7,
          height: 7,
          borderRadius: '50%',
          background: dotColor,
          flex: '0 0 auto',
        }}
      />
      {sheet.name}
    </button>
  );
}
