import { useEffect, useRef, useState } from 'react';
import { useUniverAPI } from '../use-univer';
import { useSheets } from '../hooks/useSheets';
import {
  addSheet,
  deleteSheetById,
  moveSheetTo,
  renameSheet,
  switchToSheet,
} from './sheet-actions';
import { redo, undo } from './home-tab-actions';
import { Icon } from './Icon';

/**
 * Excel-style sheet tab strip.
 *   - "+" button on the left adds a sheet.
 *   - Click a tab to switch.
 *   - Double-click a tab to rename inline (Enter commits, Esc reverts).
 *   - Hover a tab to reveal an "×" — click to delete (disabled on last sheet).
 *   - Drag a tab to reorder.
 *
 * Intentionally no right-click context menu — the X + drag pattern is
 * lighter-weight and easier to discover than a menu.
 */
export function SheetTabs() {
  const api = useUniverAPI();
  const { sheets, activeSheetId, ready } = useSheets();

  const [editingId, setEditingId] = useState<string | null>(null);
  const [draftName, setDraftName] = useState('');
  const [draggingId, setDraggingId] = useState<string | null>(null);
  const [dragOverIndex, setDragOverIndex] = useState<number | null>(null);

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

  const canDelete = sheets.length > 1;

  return (
    <div className="sheet-tabs" data-testid="sheet-tabs" role="tablist">
      <button
        type="button"
        className="sheet-tabs__add btn btn--icon"
        data-testid="sheet-tabs-add"
        aria-label="Add sheet"
        title="Add sheet"
        disabled={!ready}
        onClick={() => api && addSheet(api)}
      >
        <Icon name="add" size="sm" />
      </button>

      <div className="sheet-tabs__list" data-testid="sheet-tabs-list">
        {sheets.map((s, i) => (
          <SheetTabItem
            key={s.id}
            sheet={s}
            index={i}
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
            onDragStart={() => setDraggingId(s.id)}
            onDragOverIndex={() => setDragOverIndex(i)}
            onDragLeaveIndex={() =>
              setDragOverIndex((prev) => (prev === i ? null : prev))
            }
            onDrop={() => {
              if (api && draggingId && draggingId !== s.id) {
                moveSheetTo(api, draggingId, i);
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

      <div className="sheet-tabs__right">
        <button
          type="button"
          className="sheet-tabs__action btn btn--icon"
          data-testid="qat-undo"
          aria-label="Undo (Ctrl+Z)"
          title="Undo (Ctrl+Z)"
          disabled={!ready || !api}
          onClick={() => api && undo(api)}
        >
          <Icon name="undo" size="sm" />
        </button>
        <button
          type="button"
          className="sheet-tabs__action btn btn--icon"
          data-testid="qat-redo"
          aria-label="Redo (Ctrl+Y)"
          title="Redo (Ctrl+Y)"
          disabled={!ready || !api}
          onClick={() => api && redo(api)}
        >
          <Icon name="redo" size="sm" />
        </button>
      </div>
    </div>
  );
}

type ItemProps = {
  sheet: { id: string; name: string };
  index: number;
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
      onDoubleClick={onStartRename}
      onDragStart={(e) => {
        e.dataTransfer.effectAllowed = 'move';
        // Required for Firefox: must set some data.
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
          <button
            type="button"
            className="sheet-tab__close"
            data-testid={`sheet-tab-close-${sheet.id}`}
            aria-label={`Delete ${sheet.name}`}
            title={canDelete ? `Delete ${sheet.name}` : "Can't delete the last sheet"}
            disabled={!canDelete}
            // Stop propagation so the click doesn't also switch tabs.
            onClick={(e) => {
              e.stopPropagation();
              onDelete();
            }}
            // Don't let the drag handler eat the click.
            onMouseDown={(e) => e.stopPropagation()}
          >
            <Icon name="close" size="sm" />
          </button>
        </>
      )}
    </div>
  );
}
