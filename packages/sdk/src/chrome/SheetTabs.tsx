/**
 * SheetTabs — worksheet tab strip for `<CasualSheets chrome>`.
 *
 * The fundamental multi-sheet control the built-in chrome was missing: switch
 * sheets (click), add a sheet (+), rename (double-click), and delete (context
 * menu). Drives the editor through the FUniver facade only — `insertSheet` /
 * `setActiveSheet` / `deleteSheet` / `FWorksheet.setName` — and stays live by
 * subscribing to the sheet-lifecycle facade events (plus the mutation-level
 * fallback so collab/replay-driven changes refresh too).
 *
 * Scope of this batch: switch / add / rename / delete. Drag-reorder, tab
 * colour, hide/unhide, and duplicate are richer host features deferred to
 * later parity batches (the reference app's SheetTabs has them).
 */

import { useEffect, useState, type CSSProperties, type KeyboardEvent } from 'react';
import { createPortal } from 'react-dom';
import type { CasualSheetsAPI } from '../sheets/api';

interface SheetTab {
  id: string;
  name: string;
  hidden: boolean;
}

interface SheetsSnapshot {
  tabs: SheetTab[];
  activeId: string | null;
}

const EMPTY: SheetsSnapshot = { tabs: [], activeId: null };

/** Sheet-list mutations that must refresh the strip even when they arrive via
 *  the collab bridge's mutation-only replay (which bypasses facade events). */
const SHEET_LIST_MUTATIONS = new Set([
  'sheet.mutation.insert-sheet',
  'sheet.mutation.remove-sheet',
  'sheet.mutation.set-worksheet-name',
  'sheet.mutation.set-worksheet-order',
  'sheet.mutation.set-worksheet-hidden',
]);

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyFacade = any;

function readSheets(api: CasualSheetsAPI): SheetsSnapshot {
  const wb = api.univer.getActiveWorkbook();
  if (!wb) return EMPTY;
  const tabs = wb.getSheets().map((s) => ({
    id: s.getSheetId(),
    name: s.getSheetName(),
    hidden: (s as AnyFacade).isSheetHidden?.() === true,
  }));
  return { tabs, activeId: wb.getActiveSheet()?.getSheetId() ?? null };
}

const BAR_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'stretch',
  gap: 2,
  height: 30,
  flex: '0 0 auto',
  padding: '0 6px',
  borderTop: '1px solid var(--cs-chrome-border, rgba(0,0,0,0.12))',
  background: 'var(--cs-chrome-bg, #f8f9fa)',
  font: 'inherit',
  fontSize: 13,
  overflowX: 'auto',
  overflowY: 'hidden',
};

const ADD_BTN_STYLE: CSSProperties = {
  flex: '0 0 auto',
  width: 28,
  border: 'none',
  background: 'transparent',
  color: 'var(--cs-chrome-fg, #201f1e)',
  cursor: 'pointer',
  fontSize: 18,
  lineHeight: 1,
  borderRadius: 4,
};

const RENAME_INPUT_STYLE: CSSProperties = {
  height: 22,
  alignSelf: 'center',
  padding: '0 6px',
  border: '1px solid var(--cs-chrome-active-fg, #0e7490)',
  borderRadius: 4,
  background: 'var(--cs-chrome-input-bg, #fff)',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
  width: 120,
};

const MENU_STYLE: CSSProperties = {
  position: 'fixed',
  zIndex: 1000,
  minWidth: 140,
  padding: 4,
  border: '1px solid var(--cs-chrome-border, #e6e9ee)',
  borderRadius: 8,
  background: 'var(--cs-chrome-input-bg, #fff)',
  boxShadow: '0 6px 20px rgba(0,0,0,0.16)',
};

const MENU_ITEM_STYLE: CSSProperties = {
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

export interface SheetTabsProps {
  /** Live API, or `null` until the editor is ready. */
  api: CasualSheetsAPI | null;
}

export function SheetTabs({ api }: SheetTabsProps) {
  const [{ tabs, activeId }, setSnapshot] = useState<SheetsSnapshot>(EMPTY);
  // Inline-rename target sheet id (null = not renaming) + its draft text.
  const [renaming, setRenaming] = useState<string | null>(null);
  const [draft, setDraft] = useState('');
  // Right-click context menu: anchor + target sheet id.
  const [menu, setMenu] = useState<{ x: number; y: number; sheetId: string } | null>(null);

  useEffect(() => {
    if (!api) return;
    const refresh = () => setSnapshot(readSheets(api));
    refresh();
    const f = api.univer as AnyFacade;
    const subs: Array<{ dispose(): void }> = [];
    try {
      subs.push(
        f.addEvent(f.Event.SheetCreated, refresh),
        f.addEvent(f.Event.SheetDeleted, refresh),
        f.addEvent(f.Event.SheetNameChanged, refresh),
        f.addEvent(f.Event.ActiveSheetChanged, refresh),
        f.addEvent(f.Event.SheetMoved, refresh),
        // Mutation-level fallback — catches collab/replay-driven changes that
        // don't emit the higher-level facade events.
        f.addEvent(f.Event.CommandExecuted, (e: { id?: string }) => {
          if (e?.id && SHEET_LIST_MUTATIONS.has(e.id)) refresh();
        }),
      );
      // Whole-unit swaps (loadSnapshot / File→Open) don't fire SheetCreated;
      // refresh a tick after the new unit is wired in.
      if (typeof f.onUniverSheetCreated === 'function') {
        subs.push(f.onUniverSheetCreated(() => queueMicrotask(refresh)));
      }
    } catch {
      /* facade event surface differs — the initial read still populated tabs */
    }
    return () => {
      for (const s of subs) s?.dispose?.();
    };
  }, [api]);

  // Close the context menu on any outside interaction.
  useEffect(() => {
    if (!menu) return;
    const close = () => setMenu(null);
    window.addEventListener('pointerdown', close);
    window.addEventListener('keydown', close);
    return () => {
      window.removeEventListener('pointerdown', close);
      window.removeEventListener('keydown', close);
    };
  }, [menu]);

  if (!api) {
    return <div style={BAR_STYLE} data-testid="casual-sheets-tabs" />;
  }

  const visible = tabs.filter((t) => !t.hidden);
  const wb = () => api.univer.getActiveWorkbook();

  const switchTo = (id: string) => {
    if (id === activeId) return;
    wb()?.setActiveSheet(id);
  };

  const addSheet = () => {
    const created = (wb() as AnyFacade)?.insertSheet?.();
    created?.activate?.();
  };

  const startRename = (tab: SheetTab) => {
    setMenu(null);
    setRenaming(tab.id);
    setDraft(tab.name);
  };

  const commitRename = () => {
    const id = renaming;
    setRenaming(null);
    if (!id) return;
    const name = draft.trim();
    if (!name) return;
    const sheet = (wb() as AnyFacade)?.getSheetBySheetId?.(id);
    if (sheet && sheet.getSheetName?.() !== name) sheet.setName?.(name);
  };

  const deleteSheet = (id: string) => {
    setMenu(null);
    // Excel keeps at least one visible sheet — never delete the last one.
    if (visible.length <= 1) return;
    (wb() as AnyFacade)?.deleteSheet?.(id);
  };

  const onRenameKey = (e: KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      commitRename();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      setRenaming(null);
    }
  };

  return (
    <div style={BAR_STYLE} data-testid="casual-sheets-tabs" role="tablist">
      {visible.map((tab) => {
        const active = tab.id === activeId;
        if (renaming === tab.id) {
          return (
            <input
              key={tab.id}
              autoFocus
              style={RENAME_INPUT_STYLE}
              value={draft}
              data-testid="cs-tab-rename-input"
              onChange={(e) => setDraft(e.target.value)}
              onKeyDown={onRenameKey}
              onBlur={commitRename}
            />
          );
        }
        return (
          <button
            key={tab.id}
            type="button"
            role="tab"
            aria-selected={active}
            data-testid={`cs-tab-${tab.id}`}
            data-active={active || undefined}
            title={tab.name}
            style={{
              flex: '0 0 auto',
              maxWidth: 160,
              padding: '0 12px',
              border: 'none',
              borderTop: active
                ? '2px solid var(--cs-chrome-active-fg, #0e7490)'
                : '2px solid transparent',
              background: active ? 'var(--cs-chrome-input-bg, #fff)' : 'transparent',
              color: active
                ? 'var(--cs-chrome-active-fg, #0e7490)'
                : 'var(--cs-chrome-fg, #201f1e)',
              fontWeight: active ? 600 : 400,
              font: 'inherit',
              fontSize: 13,
              cursor: 'pointer',
              whiteSpace: 'nowrap',
              overflow: 'hidden',
              textOverflow: 'ellipsis',
            }}
            // mousedown keeps the grid from stealing focus mid-switch.
            onMouseDown={(e) => {
              e.preventDefault();
              switchTo(tab.id);
            }}
            onDoubleClick={() => startRename(tab)}
            onContextMenu={(e) => {
              e.preventDefault();
              setMenu({ x: e.clientX, y: e.clientY, sheetId: tab.id });
            }}
          >
            {tab.name}
          </button>
        );
      })}
      <button
        type="button"
        style={ADD_BTN_STYLE}
        aria-label="Add sheet"
        title="Add sheet"
        data-testid="cs-tab-add"
        onMouseDown={(e) => {
          e.preventDefault();
          addSheet();
        }}
      >
        +
      </button>

      {menu &&
        createPortal(
          <div
            // Portaled to <body> so it escapes the tab strip's stacking context
            // (otherwise the fixed menu renders BEHIND the status bar below the
            // tabs). `translateY(-100%)` opens it UPWARD from the click since the
            // tab strip sits at the bottom of the editor.
            style={{ ...MENU_STYLE, left: menu.x, top: menu.y, transform: 'translateY(-100%)' }}
            data-testid="cs-tab-menu"
            role="menu"
            // Keep clicks inside the menu from closing it via the window handler.
            onPointerDown={(e) => e.stopPropagation()}
          >
            <button
              type="button"
              role="menuitem"
              style={MENU_ITEM_STYLE}
              data-testid="cs-tab-menu-rename"
              onMouseDown={(e) => {
                e.preventDefault();
                const tab = tabs.find((t) => t.id === menu.sheetId);
                if (tab) startRename(tab);
              }}
            >
              Rename
            </button>
            <button
              type="button"
              role="menuitem"
              style={{
                ...MENU_ITEM_STYLE,
                opacity: visible.length <= 1 ? 0.5 : 1,
                cursor: visible.length <= 1 ? 'not-allowed' : 'pointer',
              }}
              data-testid="cs-tab-menu-delete"
              onMouseDown={(e) => {
                e.preventDefault();
                deleteSheet(menu.sheetId);
              }}
            >
              Delete
            </button>
          </div>,
          document.body,
        )}
    </div>
  );
}
