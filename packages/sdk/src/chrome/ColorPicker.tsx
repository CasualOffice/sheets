/**
 * ColorPicker — text-colour + fill-colour swatch pickers for `<CasualSheets chrome>`.
 *
 * Two icon buttons (text colour, fill colour). Clicking one toggles a swatch
 * popover; picking a swatch dispatches the matching Univer command on the active
 * selection and closes the popover. A "None" entry resets the colour.
 *
 * Drives the editor purely through `CasualSheetsAPI.executeCommand`. Commands:
 *   text colour : `sheet.command.set-range-text-color`  { value: '#rrggbb' | null }
 *   fill colour : `sheet.command.set-background-color`   { value: '#rrggbb' }
 *   fill reset  : `sheet.command.reset-background-color` (no params)
 * (`set-background-color` rejects a null value, so reset uses its own command.)
 */

import { useEffect, useRef, useState, type CSSProperties } from 'react';
import type { CasualSheetsAPI } from '../sheets/api';
import { Icon } from './Icon';

type Kind = 'text' | 'fill';

// A tasteful 14-swatch palette: neutrals + a spread of hues, teal = brand.
const SWATCHES: { hex: string; label: string }[] = [
  { hex: '#000000', label: 'Black' },
  { hex: '#5f6368', label: 'Dark gray' },
  { hex: '#9aa0a6', label: 'Gray' },
  { hex: '#d9dce0', label: 'Light gray' },
  { hex: '#ffffff', label: 'White' },
  { hex: '#d13438', label: 'Red' },
  { hex: '#e8710a', label: 'Orange' },
  { hex: '#f2c811', label: 'Yellow' },
  { hex: '#107c10', label: 'Green' },
  { hex: '#0e7490', label: 'Teal' },
  { hex: '#1a73e8', label: 'Blue' },
  { hex: '#5b2a86', label: 'Purple' },
  { hex: '#c2185b', label: 'Magenta' },
  { hex: '#795548', label: 'Brown' },
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
  padding: 8,
  borderRadius: 8,
  border: '1px solid var(--cs-chrome-border, #e6e9ee)',
  background: 'var(--cs-chrome-input-bg, #fff)',
  boxShadow: '0 6px 20px rgba(0,0,0,0.18)',
};

const GRID_STYLE: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: 'repeat(7, 22px)',
  gap: 6,
};

const SWATCH_STYLE: CSSProperties = {
  width: 22,
  height: 22,
  borderRadius: 4,
  border: '1px solid var(--cs-chrome-border, rgba(0,0,0,0.18))',
  cursor: 'pointer',
  padding: 0,
};

const NONE_STYLE: CSSProperties = {
  marginTop: 8,
  width: '100%',
  height: 26,
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  gap: 4,
  border: '1px solid var(--cs-chrome-border, rgba(0,0,0,0.18))',
  borderRadius: 6,
  background: 'var(--cs-chrome-input-bg, #fff)',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 12,
  cursor: 'pointer',
};

export interface ColorPickerProps {
  /** Live API, or `null` until the editor is ready. */
  api: CasualSheetsAPI | null;
}

export function ColorPicker({ api }: ColorPickerProps) {
  const [open, setOpen] = useState<Kind | null>(null);
  const rootRef = useRef<HTMLDivElement>(null);

  // Close on Escape and on any pointerdown outside the picker root.
  useEffect(() => {
    if (!open) return;
    const onPointerDown = (e: PointerEvent) => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) setOpen(null);
    };
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(null);
    };
    document.addEventListener('pointerdown', onPointerDown, true);
    document.addEventListener('keydown', onKeyDown);
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true);
      document.removeEventListener('keydown', onKeyDown);
    };
  }, [open]);

  const pick = (kind: Kind, hex: string | null) => {
    if (api) {
      if (kind === 'text') {
        void api.executeCommand('sheet.command.set-range-text-color', { value: hex });
      } else if (hex === null) {
        void api.executeCommand('sheet.command.reset-background-color');
      } else {
        void api.executeCommand('sheet.command.set-background-color', { value: hex });
      }
    }
    setOpen(null);
  };

  const toggle = (kind: Kind) => setOpen((cur) => (cur === kind ? null : kind));

  const renderButton = (kind: Kind, icon: string, label: string, testid: string) => {
    const on = open === kind;
    const baseBg = on ? 'var(--cs-chrome-active, #e6f3f7)' : 'transparent';
    return (
      <button
        type="button"
        title={label}
        aria-label={label}
        aria-haspopup="true"
        aria-expanded={on}
        data-testid={testid}
        style={{
          ...BTN_STYLE,
          background: baseBg,
          color: on ? 'var(--cs-chrome-active-fg, #0e7490)' : BTN_STYLE.color,
        }}
        // mousedown (not click) so the grid's selection isn't lost first.
        onMouseDown={(e) => {
          e.preventDefault();
          toggle(kind);
        }}
        onMouseEnter={(e) => {
          if (!on) e.currentTarget.style.background = 'var(--cs-chrome-hover, rgba(0,0,0,0.06))';
        }}
        onMouseLeave={(e) => {
          e.currentTarget.style.background = baseBg;
        }}
      >
        <Icon name={icon} size={20} />
      </button>
    );
  };

  const renderPopover = (kind: Kind) => (
    <div style={POPOVER_STYLE} data-testid="cs-color-popover" role="menu">
      <div style={GRID_STYLE}>
        {SWATCHES.map((s) => (
          <button
            key={s.hex}
            type="button"
            role="menuitem"
            title={s.label}
            aria-label={s.label}
            data-color={s.hex}
            style={{ ...SWATCH_STYLE, background: s.hex }}
            onMouseDown={(e) => {
              e.preventDefault();
              pick(kind, s.hex);
            }}
          />
        ))}
      </div>
      <button
        type="button"
        role="menuitem"
        data-color="none"
        style={NONE_STYLE}
        onMouseDown={(e) => {
          e.preventDefault();
          pick(kind, null);
        }}
      >
        <Icon name="format_color_reset" size={16} />
        None
      </button>
    </div>
  );

  return (
    <div ref={rootRef} style={ROW_STYLE} data-testid="cs-color-picker">
      {renderButton('text', 'format_color_text', 'Text color', 'cs-color-text')}
      {renderButton('fill', 'format_color_fill', 'Fill color', 'cs-color-fill')}
      {open && renderPopover(open)}
    </div>
  );
}
