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

import { useEffect, useRef, useState } from 'react';
import { useUniverAPI } from '../use-univer';
import { useSheets, type SheetSummary } from '../hooks/useSheets';
import {
  addSheet,
  deleteSheetById,
  duplicateSheet,
  hideSheet,
  moveSheetTo,
  renameSheet,
  showSheet,
  switchToSheet,
} from './sheet-actions';
import { undo } from './home-tab-actions';
import { Icon } from './Icon';
import { Tooltip } from './Tooltip';
import { useToast } from './toast/toast-context';

/**
 * Sheet-tabs strip — the tabs list (hover-X delete + drag reorder), add
 * button, and hidden-sheets menu. Selection stats, undo/redo and zoom now
 * live in the dedicated StatusBar strip below (so many tabs can scroll
 * without crowding the stats).
 */
export function SheetTabs() {
  const api = useUniverAPI();
  const toast = useToast();
  const { sheets, activeSheetId, ready } = useSheets();

  const [editingId, setEditingId] = useState<string | null>(null);
  const [draftName, setDraftName] = useState('');
  const [draggingId, setDraggingId] = useState<string | null>(null);
  const [dragOverIndex, setDragOverIndex] = useState<number | null>(null);
  const [tabMenu, setTabMenu] = useState<{ sheetId: string; x: number; y: number } | null>(null);

  const startRename = (id: string, currentName: string) => {
    setEditingId(id);
    setDraftName(currentName);
  };
  const commitRename = () => {
    if (!api || !editingId) return;
    const trimmed = draftName.trim();
    const previous = sheets.find((s) => s.id === editingId)?.name ?? '';
    const ok = renameSheet(api, editingId, trimmed);
    if (ok && trimmed !== previous) {
      toast.success(`Renamed to ${trimmed.slice(0, 31)}`);
    } else if (!ok && trimmed) {
      // renameSheet returns false on empty / duplicate / missing
      // sheet. Empty cancels silently (no-op edit); a duplicate
      // gets an explicit toast since the rename appears to succeed
      // visually but the name doesn't actually change.
      toast.error(`Couldn't rename — "${trimmed.slice(0, 31)}" is already used`);
    }
    setEditingId(null);
    setDraftName('');
  };
  const cancelRename = () => {
    setEditingId(null);
    setDraftName('');
  };

  // Delete-with-Undo — captures the sheet name BEFORE the delete
  // command runs so the toast can show what's gone, and offers an
  // Undo action button that fires Univer's command-stack undo.
  // Univer's RemoveSheetCommand pushes a redo/undo pair on the
  // history stack, so the undo recovers the deleted sheet + all
  // its data. Gmail / Google Sheets canonical pattern.
  const handleDelete = (sheetId: string) => {
    if (!api) return;
    const name = sheets.find((s) => s.id === sheetId)?.name ?? 'Sheet';
    const ok = deleteSheetById(api, sheetId);
    if (!ok) {
      toast.error("Can't delete the only sheet");
      return;
    }
    toast.success(`Deleted ${name}`, {
      action: { label: 'Undo', onClick: () => api && undo(api) },
      duration: 8000, // longer than the default so the user has
      // time to read + click Undo before it auto-dismisses.
    });
  };

  const handleDuplicate = (sheetId: string) => {
    if (!api) return;
    const name = sheets.find((s) => s.id === sheetId)?.name ?? 'Sheet';
    duplicateSheet(api, sheetId);
    // Univer mints the new sheet name asynchronously ("Sheet1
    // (2)"), so we can't tell the user the new name in this same
    // tick. Confirm the action with the SOURCE name — the new tab
    // appears in the strip in the next render.
    toast.success(`Duplicated ${name}`);
  };

  // Hide-with-Show — sheet vanishes from the tab strip but stays in
  // the workbook (formulas still resolve). The recovery path
  // (Hidden Sheets submenu) is discoverable but not obvious, so
  // offer Show as a one-click action on the toast.
  const handleHide = (sheetId: string) => {
    if (!api) return;
    const name = sheets.find((s) => s.id === sheetId)?.name ?? 'Sheet';
    const ok = hideSheet(api, sheetId);
    if (!ok) {
      toast.error("Can't hide the last visible sheet");
      return;
    }
    toast.info(`Hid ${name}`, {
      action: { label: 'Show', onClick: () => api && showSheet(api, sheetId) },
      duration: 8000,
    });
  };

  const visibleSheets = sheets.filter((s) => !s.hidden);
  const hiddenSheets = sheets.filter((s) => s.hidden);
  const canDelete = visibleSheets.length > 1;
  const canHide = visibleSheets.length > 1;
  const [hiddenMenuOpen, setHiddenMenuOpen] = useState(false);

  return (
    <div className="sheet-tabs" data-testid="sheet-tabs" role="tablist" aria-label="Sheet tabs">
      <Tooltip label="Add sheet" side="top">
        <button
          type="button"
          className="sheet-tabs__add btn btn--icon"
          data-testid="sheet-tabs-add"
          aria-label="Add sheet"
          disabled={!ready}
          onClick={() => api && addSheet(api)}
        >
          <Icon name="add" size="sm" />
        </button>
      </Tooltip>

      {hiddenSheets.length > 0 && (
        <Tooltip
          label={`${hiddenSheets.length} hidden sheet${hiddenSheets.length === 1 ? '' : 's'} — click to unhide`}
          side="top"
        >
          <button
            type="button"
            className="sheet-tabs__hidden btn btn--icon"
            data-testid="sheet-tabs-hidden"
            aria-label={`Show hidden sheets (${hiddenSheets.length})`}
            onClick={() => setHiddenMenuOpen((v) => !v)}
          >
            <Icon name="visibility_off" size="sm" />
            <span className="sheet-tabs__hidden-badge">{hiddenSheets.length}</span>
          </button>
        </Tooltip>
      )}

      <div className="sheet-tabs__list" data-testid="sheet-tabs-list">
        {visibleSheets.map((s, i) => (
          <SheetTabItem
            key={s.id}
            sheet={s}
            active={s.id === activeSheetId}
            editing={editingId === s.id}
            draftName={draftName}
            canDelete={canDelete}
            isDragging={draggingId === s.id}
            isDropTarget={dragOverIndex === i && draggingId !== null && draggingId !== s.id}
            onSwitch={() => api && switchToSheet(api, s.id)}
            onStartRename={() => startRename(s.id, s.name)}
            onDraftChange={setDraftName}
            onCommit={commitRename}
            onCancel={cancelRename}
            onDelete={() => handleDelete(s.id)}
            onContextMenu={(e) => {
              e.preventDefault();
              setTabMenu({ sheetId: s.id, x: e.clientX, y: e.clientY });
            }}
            onDragStart={() => setDraggingId(s.id)}
            onDragOverIndex={() => setDragOverIndex(i)}
            onDragLeaveIndex={() => setDragOverIndex((prev) => (prev === i ? null : prev))}
            onDrop={() => {
              if (api && draggingId && draggingId !== s.id) {
                // Translate the visible-row index to a global sheet
                // index — moveSheet operates on the full sheet order,
                // which includes hidden sheets. Without this, dropping
                // onto a visible tab in a workbook with hidden sheets
                // would land at the wrong global position.
                const globalIdx = sheets.findIndex((x) => x.id === s.id);
                if (globalIdx >= 0) moveSheetTo(api, draggingId, globalIdx);
              }
              setDraggingId(null);
              setDragOverIndex(null);
            }}
            onDragEnd={() => {
              setDraggingId(null);
              setDragOverIndex(null);
            }}
          />
        ))}
      </div>

      {tabMenu && (
        <TabContextMenu
          x={tabMenu.x}
          y={tabMenu.y}
          canDelete={canDelete}
          canHide={canHide}
          onClose={() => setTabMenu(null)}
          onRename={() => {
            const target = sheets.find((s) => s.id === tabMenu.sheetId);
            if (target) startRename(target.id, target.name);
          }}
          onDuplicate={() => handleDuplicate(tabMenu.sheetId)}
          onHide={() => handleHide(tabMenu.sheetId)}
          onDelete={() => handleDelete(tabMenu.sheetId)}
        />
      )}

      {hiddenMenuOpen && hiddenSheets.length > 0 && (
        <HiddenSheetsMenu
          sheets={hiddenSheets}
          onClose={() => setHiddenMenuOpen(false)}
          onShow={(id) => {
            if (api) showSheet(api, id);
          }}
        />
      )}
    </div>
  );
}

function HiddenSheetsMenu({
  sheets,
  onClose,
  onShow,
}: {
  sheets: SheetSummary[];
  onClose: () => void;
  onShow: (id: string) => void;
}) {
  const popRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const onDoc = (e: PointerEvent) => {
      // Don't close when the click is on the button that opened us —
      // that button toggles the same state and would re-open.
      const target = e.target as Node;
      if (popRef.current?.contains(target)) return;
      const opener = document.querySelector('[data-testid="sheet-tabs-hidden"]');
      if (opener?.contains(target)) return;
      onClose();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    document.addEventListener('pointerdown', onDoc, true);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('pointerdown', onDoc, true);
      document.removeEventListener('keydown', onKey);
    };
  }, [onClose]);

  return (
    <div
      ref={popRef}
      className="menu sheet-tabs__hidden-menu"
      role="menu"
      data-testid="hidden-sheets-menu"
    >
      <div className="menu__label">Hidden sheets</div>
      {sheets.map((s) => (
        <button
          key={s.id}
          type="button"
          role="menuitem"
          className="menu__item"
          data-testid={`hidden-sheets-menu-show-${s.id}`}
          onClick={() => {
            onShow(s.id);
            onClose();
          }}
        >
          <Icon name="visibility" size="sm" className="menu__item-icon" />
          <span>{s.name}</span>
        </button>
      ))}
    </div>
  );
}

function TabContextMenu({
  x,
  y,
  canDelete,
  canHide,
  onClose,
  onRename,
  onDuplicate,
  onHide,
  onDelete,
}: {
  x: number;
  y: number;
  canDelete: boolean;
  canHide: boolean;
  onClose: () => void;
  onRename: () => void;
  onDuplicate: () => void;
  onHide: () => void;
  onDelete: () => void;
}) {
  const popRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const onDocPointerDown = (e: PointerEvent) => {
      if (!popRef.current?.contains(e.target as Node)) onClose();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    document.addEventListener('pointerdown', onDocPointerDown, true);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('pointerdown', onDocPointerDown, true);
      document.removeEventListener('keydown', onKey);
    };
  }, [onClose]);

  // Clamp to viewport so a right-click near the bottom edge doesn't render
  // the menu off-screen.
  const top = Math.min(y, window.innerHeight - 160);
  const left = Math.min(x, window.innerWidth - 200);

  return (
    <div
      ref={popRef}
      className="menu"
      role="menu"
      data-testid="sheet-tab-context-menu"
      style={{ position: 'fixed', top, left }}
      onClick={(e) => e.stopPropagation()}
    >
      <button
        type="button"
        role="menuitem"
        className="menu__item"
        data-testid="sheet-tab-menu-rename"
        onClick={() => {
          onRename();
          onClose();
        }}
      >
        <Icon name="edit" size="sm" className="menu__item-icon" />
        <span>Rename</span>
      </button>
      <button
        type="button"
        role="menuitem"
        className="menu__item"
        data-testid="sheet-tab-menu-duplicate"
        onClick={() => {
          onDuplicate();
          onClose();
        }}
      >
        <Icon name="content_copy" size="sm" className="menu__item-icon" />
        <span>Duplicate</span>
      </button>
      <button
        type="button"
        role="menuitem"
        className="menu__item"
        data-testid="sheet-tab-menu-hide"
        disabled={!canHide}
        onClick={() => {
          onHide();
          onClose();
        }}
      >
        <Icon name="visibility_off" size="sm" className="menu__item-icon" />
        <span>Hide sheet</span>
      </button>
      <div className="menu__divider" />
      <button
        type="button"
        role="menuitem"
        className="menu__item"
        data-testid="sheet-tab-menu-delete"
        disabled={!canDelete}
        onClick={() => {
          onDelete();
          onClose();
        }}
      >
        <Icon name="delete" size="sm" className="menu__item-icon" />
        <span>Delete</span>
      </button>
    </div>
  );
}

type ItemProps = {
  sheet: { id: string; name: string };
  active: boolean;
  editing: boolean;
  draftName: string;
  canDelete: boolean;
  isDragging: boolean;
  isDropTarget: boolean;
  onSwitch: () => void;
  onStartRename: () => void;
  onDraftChange: (v: string) => void;
  onCommit: () => void;
  onCancel: () => void;
  onDelete: () => void;
  onContextMenu: (e: { preventDefault: () => void; clientX: number; clientY: number }) => void;
  onDragStart: () => void;
  onDragOverIndex: () => void;
  onDragLeaveIndex: () => void;
  onDrop: () => void;
  onDragEnd: () => void;
};

function SheetTabItem({
  sheet,
  active,
  editing,
  draftName,
  canDelete,
  isDragging,
  isDropTarget,
  onSwitch,
  onStartRename,
  onDraftChange,
  onCommit,
  onCancel,
  onDelete,
  onContextMenu,
  onDragStart,
  onDragOverIndex,
  onDragLeaveIndex,
  onDrop,
  onDragEnd,
}: ItemProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  useEffect(() => {
    if (editing) inputRef.current?.select();
  }, [editing]);

  const className = [
    'sheet-tab',
    active && 'sheet-tab--active',
    isDragging && 'sheet-tab--dragging',
    isDropTarget && 'sheet-tab--drop-target',
  ]
    .filter(Boolean)
    .join(' ');

  return (
    <div
      role="tab"
      aria-selected={active}
      // A role="tab" must be keyboard-reachable. The active tab stays in the
      // tab order; inactive tabs are reachable via roving focus (Enter/Space
      // switches, F2 renames) without flooding the Tab sequence with every
      // sheet. While editing, focus belongs to the inner <input>.
      tabIndex={editing ? -1 : active ? 0 : -1}
      className={className}
      data-testid={`sheet-tab-${sheet.id}`}
      draggable={!editing}
      onClick={() => !editing && onSwitch()}
      onKeyDown={(e) => {
        if (editing) return;
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          onSwitch();
        } else if (e.key === 'F2') {
          e.preventDefault();
          onStartRename();
        }
      }}
      onContextMenu={onContextMenu}
      onDoubleClick={onStartRename}
      onDragStart={(e) => {
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', sheet.id);
        onDragStart();
      }}
      onDragOver={(e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        onDragOverIndex();
      }}
      onDragLeave={onDragLeaveIndex}
      onDrop={(e) => {
        e.preventDefault();
        onDrop();
      }}
      onDragEnd={onDragEnd}
    >
      {editing ? (
        <input
          ref={inputRef}
          type="text"
          className="sheet-tab__input"
          data-testid={`sheet-tab-input-${sheet.id}`}
          aria-label="Sheet name"
          value={draftName}
          maxLength={31}
          onChange={(e) => onDraftChange(e.target.value)}
          onBlur={onCommit}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              onCommit();
            } else if (e.key === 'Escape') {
              e.preventDefault();
              onCancel();
            }
          }}
        />
      ) : (
        <>
          <span className="sheet-tab__label">{sheet.name}</span>
          <Tooltip
            label={canDelete ? `Delete ${sheet.name}` : "Can't delete the last sheet"}
            side="top"
          >
            <button
              type="button"
              className="sheet-tab__close"
              data-testid={`sheet-tab-close-${sheet.id}`}
              aria-label={`Delete ${sheet.name}`}
              disabled={!canDelete}
              onClick={(e) => {
                e.stopPropagation();
                onDelete();
              }}
              onMouseDown={(e) => e.stopPropagation()}
            >
              <Icon name="close" size="sm" />
            </button>
          </Tooltip>
        </>
      )}
    </div>
  );
}
