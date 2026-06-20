/**
 * AutoSumPicker — Excel-style AutoSum dropdown for `<CasualSheets chrome>`.
 *
 * One toolbar button (Σ) opening Sum / Average / Count / Max / Min. Picking one
 * inserts `=FN(<selection>)` one row below a multi-cell selection (and activates
 * that cell), or `=FN()` into a single active cell so the user can type the
 * range. Pure facade — no command, no Univer UI dependency — so it works in the
 * SDK's embedded mount.
 */

import { useEffect, useRef, useState, type CSSProperties } from 'react';
import type { CasualSheetsAPI } from '../sheets/api';
import { Icon } from './Icon';

const FUNCS: { fn: string; label: string }[] = [
  { fn: 'SUM', label: 'Sum' },
  { fn: 'AVERAGE', label: 'Average' },
  { fn: 'COUNT', label: 'Count numbers' },
  { fn: 'MAX', label: 'Max' },
  { fn: 'MIN', label: 'Min' },
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
  minWidth: 160,
  borderRadius: 8,
  border: '1px solid var(--cs-chrome-border, #e6e9ee)',
  background: 'var(--cs-chrome-input-bg, #fff)',
  boxShadow: '0 6px 20px rgba(0,0,0,0.18)',
};

const ITEM_STYLE: CSSProperties = {
  display: 'block',
  width: '100%',
  textAlign: 'left',
  height: 28,
  padding: '0 10px',
  border: 'none',
  borderRadius: 6,
  background: 'transparent',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
  cursor: 'pointer',
};

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyRange = any;

export interface AutoSumPickerProps {
  /** Live API, or `null` until the editor is ready. */
  api: CasualSheetsAPI | null;
}

export function AutoSumPicker({ api }: AutoSumPickerProps) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);

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

  const apply = (fn: string) => {
    setOpen(false);
    const sheet = api?.univer.getActiveWorkbook()?.getActiveSheet();
    const range = (sheet as unknown as { getActiveRange?: () => AnyRange })?.getActiveRange?.();
    if (!sheet || !range) return;
    const multi = range.getWidth() * range.getHeight() > 1;
    if (multi) {
      // Drop the formula one row below the selection, same column as its start.
      const target = sheet.getRange(range.getRow() + range.getHeight(), range.getColumn());
      target.setValue({ f: `=${fn}(${range.getA1Notation()})` });
      (target as AnyRange).activate?.();
    } else {
      // Single cell — leave the cursor inside the parens for the user to fill.
      sheet.getRange(range.getRow(), range.getColumn()).setValue({ f: `=${fn}()` });
    }
  };

  const on = open;
  const baseBg = on ? 'var(--cs-chrome-active, #e6f3f7)' : 'transparent';

  return (
    <div ref={rootRef} style={ROW_STYLE} data-testid="cs-autosum">
      <button
        type="button"
        title="AutoSum"
        aria-label="AutoSum"
        aria-haspopup="true"
        aria-expanded={on}
        data-testid="cs-autosum-button"
        disabled={!api}
        style={{
          ...BTN_STYLE,
          background: baseBg,
          color: on ? 'var(--cs-chrome-active-fg, #0e7490)' : BTN_STYLE.color,
        }}
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
        <Icon name="functions" size={20} />
      </button>
      {open && (
        <div style={POPOVER_STYLE} data-testid="cs-autosum-popover" role="menu">
          {FUNCS.map((f) => (
            <button
              key={f.fn}
              type="button"
              role="menuitem"
              data-fn={f.fn}
              data-testid={`cs-autosum-${f.fn}`}
              style={ITEM_STYLE}
              onMouseEnter={(e) => {
                e.currentTarget.style.background = 'var(--cs-chrome-hover, rgba(0,0,0,0.06))';
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.background = 'transparent';
              }}
              onMouseDown={(e) => {
                e.preventDefault();
                apply(f.fn);
              }}
            >
              {f.label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
