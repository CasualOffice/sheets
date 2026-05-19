import { useEffect, useRef, useState } from 'react';
import { useUniverAPI } from '../use-univer';
import { useSheets, type SheetSummary } from '../hooks/useSheets';
import { useActiveCellState } from '../hooks/useActiveCellState';
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
import { redo, undo } from './home-tab-actions';
import { setZoom } from './tab-actions';
import { Icon } from './Icon';
import { Tooltip } from './Tooltip';
import { CollabIndicator } from './CollabIndicator';

const NUM = new Intl.NumberFormat(undefined, { maximumFractionDigits: 4 });
const ZOOM_STEPS = [25, 50, 75, 100, 125, 150, 200, 300, 400];

/**
 * Bottom strip — combines:
 *   • the sheet-tabs list with hover-X delete + drag reorder + add button
 *   • selection stats (count / sum / avg, multi-cell only)
 *   • undo / redo
 *   • zoom slider with − / + steppers and a click-to-reset 100%% label
 *
 * Replaces the previous separate status bar. Status info that isn't
 * actionable (e.g. "Ready") was dropped — empty signal isn't worth a row.
 */
export function SheetTabs() {
  const api = useUniverAPI();
  const { sheets, activeSheetId, ready } = useSheets();
  const { stats } = useActiveCellState();

  const [editingId, setEditingId] = useState<string | null>(null);
  const [draftName, setDraftName] = useState('');
  const [draggingId, setDraggingId] = useState<string | null>(null);
  const [dragOverIndex, setDragOverIndex] = useState<number | null>(null);
  const [tabMenu, setTabMenu] = useState<{ sheetId: string; x: number; y: number } | null>(null);

  // Zoom — synced from set-zoom-ratio command params, accurate.
  const [zoomPct, setZoomPct] = useState(100);
  useEffect(() => {
    if (!api) return;
    const d = api.addEvent(api.Event.CommandExecuted, (e) => {
      const info = e as { id?: string; params?: { zoomRatio?: number } };
      if (info.id === 'sheet.command.set-zoom-ratio' && typeof info.params?.zoomRatio === 'number') {
        setZoomPct(Math.round(info.params.zoomRatio * 100));
      }
    });
    return () => d.dispose();
  }, [api]);

  const applyZoom = (pct: number) => {
    if (!api) return;
    const clamped = Math.max(25, Math.min(400, pct));
    setZoomPct(clamped);
    setZoom(api, clamped / 100);
  };

  const startRename = (id: string, currentName: string) => {
    setEditingId(id);
    setDraftName(currentName);
  };
  const commitRename = () => {
    if (!api || !editingId) return;
    renameSheet(api, editingId, draftName.trim());
    setEditingId(null);
    setDraftName('');
  };
  const cancelRename = () => {
    setEditingId(null);
    setDraftName('');
  };

  const visibleSheets = sheets.filter((s) => !s.hidden);
  const hiddenSheets = sheets.filter((s) => s.hidden);
  const canDelete = visibleSheets.length > 1;
  const canHide = visibleSheets.length > 1;
  const [hiddenMenuOpen, setHiddenMenuOpen] = useState(false);

  return (
    <div className="sheet-tabs" data-testid="sheet-tabs" role="tablist">
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
            onDelete={() => api && deleteSheetById(api, s.id)}
            onContextMenu={(e) => {
              e.preventDefault();
              setTabMenu({ sheetId: s.id, x: e.clientX, y: e.clientY });
            }}
            onDragStart={() => setDraggingId(s.id)}
            onDragOverIndex={() => setDragOverIndex(i)}
            onDragLeaveIndex={() =>
              setDragOverIndex((prev) => (prev === i ? null : prev))
            }
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

      {stats && stats.cellCount > 0 && (
        <div className="sheet-tabs__stats" data-testid="sheet-tabs-stats">
          {/* Excel order: Average, Count, Numerical Count, Min, Max, Sum.
              We show Avg / Min / Max only when at least one numeric value
              participated; Count / Numerical Count / Sum always render
              when any cell is selected. */}
          {stats.avg !== null && (
            <span data-testid="stat-avg">Average: {NUM.format(stats.avg)}</span>
          )}
          <span data-testid="stat-count">Count: {stats.cellCount}</span>
          {stats.count !== stats.cellCount && (
            <span data-testid="stat-num-count">Numerical Count: {stats.count}</span>
          )}
          {stats.min !== null && (
            <span data-testid="stat-min">Min: {NUM.format(stats.min)}</span>
          )}
          {stats.max !== null && (
            <span data-testid="stat-max">Max: {NUM.format(stats.max)}</span>
          )}
          <span data-testid="stat-sum">Sum: {NUM.format(stats.sum)}</span>
        </div>
      )}

      <div className="sheet-tabs__right">
        <CollabIndicator />
        <span className="sheet-tabs__sep" aria-hidden="true" />
        <Tooltip label="Undo (Ctrl+Z)" side="top">
          <button
            type="button"
            className="sheet-tabs__action btn btn--icon"
            data-testid="qat-undo"
            aria-label="Undo (Ctrl+Z)"
            disabled={!ready || !api}
            onClick={() => api && undo(api)}
          >
            <Icon name="undo" size="sm" />
          </button>
        </Tooltip>
        <Tooltip label="Redo (Ctrl+Y)" side="top">
          <button
            type="button"
            className="sheet-tabs__action btn btn--icon"
            data-testid="qat-redo"
            aria-label="Redo (Ctrl+Y)"
            disabled={!ready || !api}
            onClick={() => api && redo(api)}
          >
            <Icon name="redo" size="sm" />
          </button>
        </Tooltip>

        <span className="sheet-tabs__sep" aria-hidden="true" />

        <Tooltip label="Zoom out" side="top">
          <button
            type="button"
            className="sheet-tabs__action btn btn--icon"
            data-testid="statusbar-zoom-out"
            aria-label="Zoom out"
            onClick={() => {
              const prev = [...ZOOM_STEPS].reverse().find((s) => s < zoomPct);
              applyZoom(prev ?? 25);
            }}
          >
            <Icon name="zoom_out" size="sm" />
          </button>
        </Tooltip>
        <input
          type="range"
          min={25}
          max={400}
          step={5}
          value={zoomPct}
          data-testid="statusbar-zoom-slider"
          aria-label="Zoom slider"
          className="sheet-tabs__zoom-slider"
          onChange={(e) => applyZoom(Number(e.target.value))}
        />
        <Tooltip label="Zoom in" side="top">
          <button
            type="button"
            className="sheet-tabs__action btn btn--icon"
            data-testid="statusbar-zoom-in"
            aria-label="Zoom in"
            onClick={() => {
              const next = ZOOM_STEPS.find((s) => s > zoomPct);
              applyZoom(next ?? 400);
            }}
          >
            <Icon name="zoom_in" size="sm" />
          </button>
        </Tooltip>
        <Tooltip label="Reset to 100%" side="top">
          <button
            type="button"
            className="sheet-tabs__zoom-label"
            data-testid="statusbar-zoom-label"
            aria-label="Reset zoom to 100%"
            onClick={() => applyZoom(100)}
          >
            {zoomPct}%
          </button>
        </Tooltip>
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
          onDuplicate={() => api && duplicateSheet(api, tabMenu.sheetId)}
          onHide={() => api && hideSheet(api, tabMenu.sheetId)}
          onDelete={() => api && deleteSheetById(api, tabMenu.sheetId)}
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
      className={className}
      data-testid={`sheet-tab-${sheet.id}`}
      draggable={!editing}
      onClick={() => !editing && onSwitch()}
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
