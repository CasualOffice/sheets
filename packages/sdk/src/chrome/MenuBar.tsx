/**
 * MenuBar — the built-in dropdown menu row for `<CasualSheets chrome>`.
 *
 * Office-style horizontal menu strip (no logo, no title — the host frames the
 * editor with its own bar): Edit · Insert · Format · Data. Each button opens a
 * dropdown of items that dispatch a single verified Univer command through
 * `CasualSheetsAPI.executeCommand`, then close the menu.
 *
 * Self-contained: only one menu is open at a time; Escape and an outside
 * pointerdown close it. Every command id below was verified against the fork
 * (`vendor/univer-revamp/packages`); items whose command can't be verified are
 * omitted rather than dispatched blindly.
 */

import { useEffect, useRef, useState, type CSSProperties } from 'react';
import type { CasualSheetsAPI } from '../sheets/api';
import { Icon } from './Icon';
import { ensureChromeFonts } from './fonts';

type MenuId = 'edit' | 'insert' | 'format' | 'data';

interface MenuItem {
  /** Stable id — drives the per-item testid `cs-menuitem-<id>`. */
  id: string;
  label: string;
  command: string;
  icon?: string;
  params?: object;
}

interface Menu {
  id: MenuId;
  label: string;
  items: MenuItem[];
}

// All command ids below verified via:
//   grep -rhoE "id: '(sheet\.command\.[a-z.-]+)'" vendor/univer-revamp/packages | sort -u
//   (undo/redo: univer.command.undo / .redo in packages/core undoredo)
const MENUS: Menu[] = [
  {
    id: 'edit',
    label: 'Edit',
    items: [
      { id: 'undo', label: 'Undo', command: 'univer.command.undo', icon: 'undo' },
      { id: 'redo', label: 'Redo', command: 'univer.command.redo', icon: 'redo' },
    ],
  },
  {
    id: 'insert',
    label: 'Insert',
    items: [
      {
        id: 'insert-row',
        label: 'Insert row above',
        command: 'sheet.command.insert-row-before',
        icon: 'add_row_above',
      },
      {
        id: 'insert-col',
        label: 'Insert column left',
        command: 'sheet.command.insert-col-before',
        icon: 'add_column_left',
      },
    ],
  },
  {
    id: 'format',
    label: 'Format',
    items: [
      { id: 'bold', label: 'Bold', command: 'sheet.command.set-range-bold', icon: 'format_bold' },
      {
        id: 'italic',
        label: 'Italic',
        command: 'sheet.command.set-range-italic',
        icon: 'format_italic',
      },
      {
        id: 'underline',
        label: 'Underline',
        command: 'sheet.command.set-range-underline',
        icon: 'format_underlined',
      },
    ],
  },
  {
    id: 'data',
    label: 'Data',
    items: [
      {
        id: 'sort-asc',
        label: 'Sort ascending',
        command: 'sheet.command.sort-range-asc',
        icon: 'arrow_upward',
      },
      {
        id: 'sort-desc',
        label: 'Sort descending',
        command: 'sheet.command.sort-range-desc',
        icon: 'arrow_downward',
      },
      {
        id: 'toggle-filter',
        label: 'Toggle filter',
        command: 'sheet.command.smart-toggle-filter',
        icon: 'filter_alt',
      },
    ],
  },
];

const BAR_STYLE: CSSProperties = {
  position: 'relative',
  display: 'flex',
  alignItems: 'center',
  gap: 2,
  padding: '2px 6px',
  borderBottom: '1px solid var(--cs-chrome-border, #e6e9ee)',
  background: 'var(--cs-chrome-bg, #eef1f5)',
  flex: '0 0 auto',
  userSelect: 'none',
  font: 'inherit',
  fontSize: 13,
};

const MENU_BTN_STYLE: CSSProperties = {
  height: 26,
  display: 'inline-flex',
  alignItems: 'center',
  padding: '0 10px',
  border: 'none',
  borderRadius: 6,
  background: 'transparent',
  cursor: 'pointer',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
};

const DROPDOWN_STYLE: CSSProperties = {
  position: 'absolute',
  top: '100%',
  minWidth: 200,
  marginTop: 2,
  padding: 4,
  display: 'flex',
  flexDirection: 'column',
  border: '1px solid var(--cs-chrome-border, #e6e9ee)',
  borderRadius: 8,
  background: 'var(--cs-chrome-input-bg, #fff)',
  boxShadow: '0 6px 20px rgba(0,0,0,0.16)',
  zIndex: 1000,
};

const ITEM_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 10,
  width: '100%',
  height: 30,
  padding: '0 10px',
  border: 'none',
  borderRadius: 6,
  background: 'transparent',
  cursor: 'pointer',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
  textAlign: 'left',
};

export interface MenuBarProps {
  /** Live API, or `null` until the editor is ready. */
  api: CasualSheetsAPI | null;
}

export function MenuBar({ api }: MenuBarProps) {
  const [open, setOpen] = useState<MenuId | null>(null);
  const rootRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    ensureChromeFonts();
  }, []);

  // Close on Escape + on a pointerdown outside the menu bar.
  useEffect(() => {
    if (open === null) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(null);
    };
    const onDown = (e: PointerEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) setOpen(null);
    };
    document.addEventListener('keydown', onKey);
    document.addEventListener('pointerdown', onDown, true);
    return () => {
      document.removeEventListener('keydown', onKey);
      document.removeEventListener('pointerdown', onDown, true);
    };
  }, [open]);

  const run = (item: MenuItem) => {
    setOpen(null);
    void api?.executeCommand(item.command, item.params);
  };

  return (
    <div
      ref={rootRef}
      style={BAR_STYLE}
      data-testid="cs-menubar"
      role="menubar"
      aria-label="Menu bar"
    >
      {MENUS.map((menu) => {
        const isOpen = open === menu.id;
        return (
          <div key={menu.id} style={{ position: 'relative' }}>
            <button
              type="button"
              data-menu={menu.id}
              aria-haspopup="menu"
              aria-expanded={isOpen}
              style={{
                ...MENU_BTN_STYLE,
                background: isOpen ? 'var(--cs-chrome-active, #e6f3f7)' : 'transparent',
                color: isOpen ? 'var(--cs-chrome-active-fg, #0e7490)' : MENU_BTN_STYLE.color,
              }}
              // mousedown (not click) so the grid selection isn't lost first;
              // toggle the menu open/closed.
              onMouseDown={(e) => {
                e.preventDefault();
                setOpen((cur) => (cur === menu.id ? null : menu.id));
              }}
              onMouseEnter={(e) => {
                // Hover-to-switch once a menu is already open (Office behaviour).
                if (open !== null && open !== menu.id) setOpen(menu.id);
                else if (!isOpen)
                  e.currentTarget.style.background = 'var(--cs-chrome-hover, rgba(0,0,0,0.06))';
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.background = isOpen
                  ? 'var(--cs-chrome-active, #e6f3f7)'
                  : 'transparent';
              }}
            >
              {menu.label}
            </button>
            {isOpen && (
              <div style={DROPDOWN_STYLE} role="menu" aria-label={menu.label}>
                {menu.items.map((item) => (
                  <button
                    key={item.id}
                    type="button"
                    role="menuitem"
                    data-testid={`cs-menuitem-${item.id}`}
                    disabled={!api}
                    style={{ ...ITEM_STYLE, opacity: api ? 1 : 0.5 }}
                    onMouseDown={(e) => {
                      e.preventDefault();
                      run(item);
                    }}
                    onMouseEnter={(e) => {
                      e.currentTarget.style.background = 'var(--cs-chrome-hover, rgba(0,0,0,0.06))';
                    }}
                    onMouseLeave={(e) => {
                      e.currentTarget.style.background = 'transparent';
                    }}
                  >
                    {item.icon ? (
                      <Icon name={item.icon} size={18} />
                    ) : (
                      <span style={{ width: 18 }} aria-hidden />
                    )}
                    <span>{item.label}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
