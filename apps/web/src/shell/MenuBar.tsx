import { useEffect, useRef, useState, type ReactNode } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { Icon } from './Icon';
import { PropertiesDialog } from './PropertiesDialog';
import { AboutDialog } from './AboutDialog';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { useUI } from '../use-ui';
import { emptyWorkbook } from '../snapshot';
import { openXlsx, pickXlsxFile, saveAsXlsx } from './file-actions';
import {
  copy as actCopy,
  cut as actCut,
  decreaseDecimal,
  increaseDecimal,
  openFindReplace,
  paste as actPaste,
  redo,
  setNumberFormatByKey,
  undo,
  type NumberFormatKey,
} from './home-tab-actions';
import {
  autoFitColumns,
  autoFitRows,
  deleteSelectedColumn,
  deleteSelectedRow,
  freezeAtSelection,
  freezeFirstColumn,
  freezeFirstRow,
  hideSelectedColumns,
  hideSelectedRows,
  insertColumnLeft,
  insertColumnRight,
  insertComment,
  insertHyperlink,
  insertImage,
  insertNewSheet,
  insertRowAbove,
  insertRowBelow,
  insertTable,
  openConditionalFormatting,
  openDataValidation,
  removeDuplicates,
  splitTextToColumns,
  toggleCommentPanel,
  toggleGridlines,
  unfreezePanes,
  unhideSelectedColumns,
  unhideSelectedRows,
} from './tab-actions';

/**
 * Google-Sheets-style menu bar: File / Edit / View / Insert / Format / Data / Help.
 * Each top-level button opens a dropdown of items. Clicking an item dispatches
 * a Univer command via the actions modules. Only one menu open at a time.
 */

type MenuId = 'file' | 'edit' | 'view' | 'insert' | 'format' | 'data' | 'help';

type MenuItem =
  | {
      kind: 'item';
      id: string;
      label: string;
      icon?: string;
      shortcut?: string;
      run?: (api: FUniver) => void;
      onClick?: () => void;
      disabled?: boolean;
    }
  | { kind: 'separator'; id: string };

export function MenuBar() {
  const api = useUniverAPI();
  const workbook = useWorkbook();
  const ui = useUI();
  const [open, setOpen] = useState<MenuId | null>(null);
  const [showProperties, setShowProperties] = useState(false);
  const [showAbout, setShowAbout] = useState(false);

  const onClose = () => setOpen(null);

  const handleNew = () => workbook.replaceWorkbook(emptyWorkbook());
  const handleOpen = async () => {
    const file = await pickXlsxFile();
    if (file) {
      const data = await openXlsx(file);
      workbook.replaceWorkbook(data);
    }
  };
  const handleSaveAs = async () => {
    if (!api) return;
    await saveAsXlsx(api, workbook.snapshot.name || 'workbook');
  };

  const menus: Record<MenuId, { label: string; items: MenuItem[] }> = {
    file: {
      label: 'File',
      items: [
        { kind: 'item', id: 'new', label: 'New', icon: 'add', shortcut: 'Ctrl+N', onClick: handleNew },
        { kind: 'item', id: 'open', label: 'Open', icon: 'folder_open', shortcut: 'Ctrl+O', onClick: handleOpen },
        { kind: 'item', id: 'save-as', label: 'Save As', icon: 'save', shortcut: 'Ctrl+Shift+S', onClick: handleSaveAs },
        { kind: 'separator', id: 'sep-1' },
        { kind: 'item', id: 'properties', label: 'Properties', icon: 'info', onClick: () => setShowProperties(true) },
      ],
    },
    edit: {
      label: 'Edit',
      items: [
        { kind: 'item', id: 'undo', label: 'Undo', icon: 'undo', shortcut: 'Ctrl+Z', run: undo },
        { kind: 'item', id: 'redo', label: 'Redo', icon: 'redo', shortcut: 'Ctrl+Y', run: redo },
        { kind: 'separator', id: 'sep-1' },
        { kind: 'item', id: 'cut', label: 'Cut', icon: 'content_cut', shortcut: 'Ctrl+X', run: actCut },
        { kind: 'item', id: 'copy', label: 'Copy', icon: 'content_copy', shortcut: 'Ctrl+C', run: actCopy },
        { kind: 'item', id: 'paste', label: 'Paste', icon: 'content_paste', shortcut: 'Ctrl+V', run: actPaste },
        { kind: 'separator', id: 'sep-2' },
        { kind: 'item', id: 'find-replace', label: 'Find & Replace', icon: 'search', shortcut: 'Ctrl+F', run: openFindReplace },
      ],
    },
    view: {
      label: 'View',
      items: [
        { kind: 'item', id: 'freeze-row', label: 'Freeze top row', icon: 'border_horizontal', run: freezeFirstRow },
        { kind: 'item', id: 'freeze-col', label: 'Freeze first column', icon: 'border_vertical', run: freezeFirstColumn },
        { kind: 'item', id: 'freeze-selection', label: 'Freeze panes (at selection)', icon: 'grid_4x4', run: freezeAtSelection },
        { kind: 'item', id: 'unfreeze', label: 'Unfreeze', icon: 'grid_off', run: unfreezePanes },
        { kind: 'separator', id: 'sep-1' },
        { kind: 'item', id: 'toggle-gridlines', label: 'Gridlines', icon: 'grid_on', onClick: () => api && toggleGridlines(api, true) },
        { kind: 'item', id: 'toggle-formula-bar', label: ui.formulaBarVisible ? 'Hide formula bar' : 'Show formula bar', icon: 'functions', onClick: ui.toggleFormulaBar },
      ],
    },
    insert: {
      label: 'Insert',
      items: [
        { kind: 'item', id: 'insert-row-above', label: 'Row above', icon: 'vertical_align_top', run: insertRowAbove },
        { kind: 'item', id: 'insert-row-below', label: 'Row below', icon: 'vertical_align_bottom', run: insertRowBelow },
        { kind: 'item', id: 'insert-col-left', label: 'Column left', icon: 'keyboard_tab_rtl', run: insertColumnLeft },
        { kind: 'item', id: 'insert-col-right', label: 'Column right', icon: 'keyboard_tab', run: insertColumnRight },
        { kind: 'separator', id: 'sep-1' },
        { kind: 'item', id: 'delete-row', label: 'Delete row', icon: 'delete_sweep', run: deleteSelectedRow },
        { kind: 'item', id: 'delete-col', label: 'Delete column', icon: 'folder_delete', run: deleteSelectedColumn },
        { kind: 'separator', id: 'sep-2' },
        { kind: 'item', id: 'hide-row', label: 'Hide row', icon: 'visibility_off', run: hideSelectedRows },
        { kind: 'item', id: 'unhide-row', label: 'Unhide row', icon: 'visibility', run: unhideSelectedRows },
        { kind: 'item', id: 'hide-col', label: 'Hide column', icon: 'visibility_off', run: hideSelectedColumns },
        { kind: 'item', id: 'unhide-col', label: 'Unhide column', icon: 'visibility', run: unhideSelectedColumns },
        { kind: 'separator', id: 'sep-3' },
        { kind: 'item', id: 'new-sheet', label: 'New sheet', icon: 'add_box', run: insertNewSheet },
        { kind: 'item', id: 'insert-table', label: 'Table', icon: 'table_rows', run: insertTable },
        { kind: 'item', id: 'insert-image', label: 'Image', icon: 'image', run: insertImage },
        { kind: 'item', id: 'insert-link', label: 'Hyperlink', icon: 'link', shortcut: 'Ctrl+K', run: insertHyperlink },
        { kind: 'item', id: 'insert-comment', label: 'Comment', icon: 'comment', run: insertComment },
        { kind: 'separator', id: 'sep-4' },
        { kind: 'item', id: 'autofit-col', label: 'Auto-fit column width', icon: 'settings_ethernet', run: autoFitColumns },
        { kind: 'item', id: 'autofit-row', label: 'Auto-fit row height', icon: 'height', run: autoFitRows },
      ],
    },
    format: {
      label: 'Format',
      items: [
        ...(
          ['general', 'number', 'integer', 'currency', 'accounting', 'percent', 'date', 'time', 'scientific', 'text'] as NumberFormatKey[]
        ).map<MenuItem>((k) => ({
          kind: 'item',
          id: `num-${k}`,
          label: k[0]!.toUpperCase() + k.slice(1),
          icon: 'looks_one',
          onClick: () => api && setNumberFormatByKey(api, k),
        })),
        { kind: 'separator', id: 'sep-1' },
        { kind: 'item', id: 'decimal-up', label: 'Increase decimals', icon: 'add', run: increaseDecimal },
        { kind: 'item', id: 'decimal-down', label: 'Decrease decimals', icon: 'remove', run: decreaseDecimal },
      ],
    },
    data: {
      label: 'Data',
      items: [
        { kind: 'item', id: 'data-validation', label: 'Data validation…', icon: 'rule', run: openDataValidation },
        { kind: 'item', id: 'conditional-formatting', label: 'Conditional formatting…', icon: 'palette', run: openConditionalFormatting },
        { kind: 'separator', id: 'sep-1' },
        { kind: 'item', id: 'text-to-columns', label: 'Text to Columns', icon: 'splitscreen', run: splitTextToColumns },
        { kind: 'item', id: 'remove-duplicates', label: 'Remove Duplicates', icon: 'filter_list_off', run: removeDuplicates },
        { kind: 'separator', id: 'sep-2' },
        { kind: 'item', id: 'comments-panel', label: 'Comments panel', icon: 'forum', run: toggleCommentPanel },
      ],
    },
    help: {
      label: 'Help',
      items: [
        {
          kind: 'item',
          id: 'about',
          label: 'About Casual Sheets',
          icon: 'info',
          onClick: () => setShowAbout(true),
        },
        {
          kind: 'item',
          id: 'github',
          label: 'View on GitHub',
          icon: 'open_in_new',
          onClick: () => window.open('https://github.com/schnsrw/sheets', '_blank'),
        },
      ],
    },
  };

  return (
    <>
      <div className="menubar" role="menubar" data-testid="menubar">
        {(Object.keys(menus) as MenuId[]).map((id) => (
          <MenuItemButton
            key={id}
            id={id}
            label={menus[id].label}
            isOpen={open === id}
            onToggle={() => setOpen(open === id ? null : id)}
            onHoverOpen={() => open !== null && setOpen(id)}
          >
            <MenuList
              items={menus[id].items}
              onItemClick={(item) => {
                if (item.kind !== 'item') return;
                if (item.disabled) return;
                if (item.run && api) item.run(api);
                if (item.onClick) item.onClick();
                onClose();
              }}
            />
          </MenuItemButton>
        ))}
      </div>

      {showProperties && (
        <PropertiesDialog
          onClose={() => setShowProperties(false)}
        />
      )}

      {showAbout && <AboutDialog onClose={() => setShowAbout(false)} />}
    </>
  );
}

function MenuItemButton({
  id,
  label,
  isOpen,
  onToggle,
  onHoverOpen,
  children,
}: {
  id: MenuId;
  label: string;
  isOpen: boolean;
  onToggle: () => void;
  onHoverOpen: () => void;
  children: ReactNode;
}) {
  const btnRef = useRef<HTMLButtonElement>(null);
  const popRef = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ top: number; left: number }>({ top: 0, left: 0 });

  useEffect(() => {
    if (!isOpen) return;
    const rect = btnRef.current?.getBoundingClientRect();
    if (rect) setPos({ top: rect.bottom + 2, left: rect.left });
  }, [isOpen]);

  useEffect(() => {
    if (!isOpen) return;
    const onDocPointerDown = (e: PointerEvent) => {
      if (
        !popRef.current?.contains(e.target as Node) &&
        !btnRef.current?.contains(e.target as Node)
      ) {
        onToggle();
      }
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onToggle();
    };
    document.addEventListener('pointerdown', onDocPointerDown, true);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('pointerdown', onDocPointerDown, true);
      document.removeEventListener('keydown', onKey);
    };
  }, [isOpen, onToggle]);

  return (
    <>
      <button
        ref={btnRef}
        type="button"
        role="menuitem"
        className={`menubar__item${isOpen ? ' menubar__item--open' : ''}`}
        data-testid={`menubar-${id}`}
        aria-haspopup="menu"
        aria-expanded={isOpen}
        onClick={onToggle}
        onMouseEnter={onHoverOpen}
      >
        {label}
      </button>
      {isOpen && (
        <div
          ref={popRef}
          className="menu"
          role="menu"
          data-testid={`menubar-${id}-popup`}
          style={{ top: pos.top, left: pos.left }}
          onClick={(e) => e.stopPropagation()}
        >
          {children}
        </div>
      )}
    </>
  );
}

function MenuList({
  items,
  onItemClick,
}: {
  items: MenuItem[];
  onItemClick: (item: MenuItem) => void;
}) {
  return (
    <>
      {items.map((item) =>
        item.kind === 'separator' ? (
          <div key={item.id} className="menu__divider" />
        ) : (
          <button
            key={item.id}
            type="button"
            role="menuitem"
            className="menu__item"
            data-testid={`menu-item-${item.id}`}
            disabled={item.disabled}
            onClick={() => onItemClick(item)}
          >
            {item.icon && <Icon name={item.icon} size="sm" className="menu__item-icon" />}
            <span>{item.label}</span>
            {item.shortcut && <span className="menu__item-shortcut">{item.shortcut}</span>}
          </button>
        ),
      )}
    </>
  );
}
