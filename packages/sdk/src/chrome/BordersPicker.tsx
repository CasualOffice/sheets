/**
 * BordersPicker — border-position dropdown for `<CasualSheets chrome>`.
 *
 * One toolbar button that opens a popover of Excel-style border options (All,
 * Outside, Inside, Top, Bottom, Left, Right, None). Picking one dispatches
 * `sheet.command.set-border-position` with the matching `BorderType` against the
 * active selection, using Univer's current border style/colour (thin black by
 * default). Drives the editor purely through `CasualSheetsAPI.executeCommand`.
 */

import { useEffect, useRef, useState, type CSSProperties } from 'react';
import type { CasualSheetsAPI } from '../sheets/api';
import { Icon } from './Icon';

// BorderType is a string enum in @univerjs/core ('all' | 'outside' | …); use the
// literal values directly so this component needs no Univer value import.
const OPTIONS: { value: string; icon: string; label: string }[] = [
  { value: 'all', icon: 'border_all', label: 'All borders' },
  { value: 'outside', icon: 'border_outer', label: 'Outside borders' },
  { value: 'inside', icon: 'border_inner', label: 'Inside borders' },
  { value: 'top', icon: 'border_top', label: 'Top border' },
  { value: 'bottom', icon: 'border_bottom', label: 'Bottom border' },
  { value: 'left', icon: 'border_left', label: 'Left border' },
  { value: 'right', icon: 'border_right', label: 'Right border' },
  { value: 'none', icon: 'border_clear', label: 'No border' },
];

const ROW_STYLE: CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  position: 'relative',
};

const BTN_STYLE: CSSProperties = {
  width: 30,
  height: 30,
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  border: 'none',
  borderRadius: 6,
  background: 'transparent',
  cursor: 'pointer',
  color: 'var(--cs-chrome-fg, #201f1e)',
  padding: 0,
};

const POPOVER_STYLE: CSSProperties = {
  position: 'absolute',
  top: '100%',
  left: 0,
  marginTop: 4,
  zIndex: 1000,
  padding: 4,
  minWidth: 168,
  borderRadius: 8,
  border: '1px solid var(--cs-chrome-border, #e6e9ee)',
  background: 'var(--cs-chrome-input-bg, #fff)',
  boxShadow: '0 6px 20px rgba(0,0,0,0.18)',
};

const ITEM_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  width: '100%',
  height: 28,
  padding: '0 8px',
  border: 'none',
  borderRadius: 6,
  background: 'transparent',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
  cursor: 'pointer',
  textAlign: 'left',
};

export interface BordersPickerProps {
  /** Live API, or `null` until the editor is ready. */
  api: CasualSheetsAPI | null;
}

export function BordersPicker({ api }: BordersPickerProps) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);

  // Close on Escape and on any pointerdown outside the picker root.
  useEffect(() => {
    if (!open) return;
    const onPointerDown = (e: PointerEvent) => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) setOpen(false);
    };
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(false);
    };
    document.addEventListener('pointerdown', onPointerDown, true);
    document.addEventListener('keydown', onKeyDown);
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true);
      document.removeEventListener('keydown', onKeyDown);
    };
  }, [open]);

  const pick = (value: string) => {
    void api?.executeCommand('sheet.command.set-border-position', { value });
    setOpen(false);
  };

  const on = open;
  const baseBg = on ? 'var(--cs-chrome-active, #e6f3f7)' : 'transparent';

  return (
    <div ref={rootRef} style={ROW_STYLE} data-testid="cs-borders">
      <button
        type="button"
        title="Borders"
        aria-label="Borders"
        aria-haspopup="true"
        aria-expanded={on}
        data-testid="cs-borders-button"
        disabled={!api}
        style={{
          ...BTN_STYLE,
          background: baseBg,
          color: on ? 'var(--cs-chrome-active-fg, #0e7490)' : BTN_STYLE.color,
        }}
        // mousedown (not click) so the grid's selection isn't lost first.
        onMouseDown={(e) => {
          e.preventDefault();
          setOpen((v) => !v);
        }}
        onMouseEnter={(e) => {
          if (!on) e.currentTarget.style.background = 'var(--cs-chrome-hover, rgba(0,0,0,0.06))';
        }}
        onMouseLeave={(e) => {
          e.currentTarget.style.background = baseBg;
        }}
      >
        <Icon name="border_all" size={20} />
      </button>
      {open && (
        <div style={POPOVER_STYLE} data-testid="cs-borders-popover" role="menu">
          {OPTIONS.map((o) => (
            <button
              key={o.value}
              type="button"
              role="menuitem"
              data-border={o.value}
              data-testid={`cs-border-${o.value}`}
              style={ITEM_STYLE}
              onMouseEnter={(e) => {
                e.currentTarget.style.background = 'var(--cs-chrome-hover, rgba(0,0,0,0.06))';
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.background = 'transparent';
              }}
              onMouseDown={(e) => {
                e.preventDefault();
                pick(o.value);
              }}
            >
              <Icon name={o.icon} size={18} />
              {o.label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
