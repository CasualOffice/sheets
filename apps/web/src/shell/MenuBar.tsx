import { useEffect, useRef, useState, type ReactNode } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { Icon } from './Icon';
import { PropertiesDialog } from './PropertiesDialog';
import { AboutDialog } from './AboutDialog';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { useUI } from '../use-ui';
import { emptyWorkbook } from '../snapshot';
import {
  loadSpreadsheetFile,
  pickXlsxFile,
  saveAsCsv,
  saveAsOds,
  saveAsTsv,
  saveAsXlsx,
} from './file-actions';
import { printActiveSheet } from './print';
import { openBugReport } from './report-bug';
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
  showAllRows,
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
  | { kind: 'separator'; id: string }
  | {
      kind: 'submenu';
      id: string;
      label: string;
      icon?: string;
      items: MenuItem[];
    };

export function MenuBar() {
  const api = useUniverAPI();
  const workbook = useWorkbook();
  const ui = useUI();
  const [open, setOpen] = useState<MenuId | null>(null);
  const [showProperties, setShowProperties] = useState(false);
  const [showAbout, setShowAbout] = useState(false);

  const onClose = () => setOpen(null);

  // Intercept Ctrl/Cmd+P globally — the default would print the whole web
  // page (chrome + grid). We print only the active sheet via an offscreen
  // iframe instead. Capture-phase so we beat browser shortcuts on focused
  // inputs as well. Same hook owns Ctrl/Cmd+S → Save (in source format).
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const mod = e.metaKey || e.ctrlKey;
      const k = e.key.toLowerCase();
      if (mod && !e.altKey) {
        if (k === 'p' && !e.shiftKey) {
          e.preventDefault();
          if (api) printActiveSheet(api);
        } else if (k === 's' && !e.shiftKey) {
          e.preventDefault();
          void handleSave();
        }
      }
    };
    window.addEventListener('keydown', onKey, { capture: true });
    return () => window.removeEventListener('keydown', onKey, { capture: true });
    // handleSave reads workbook.sourceFormat + workbook.snapshot.name from
    // context; api in deps re-binds the handler when the workbook is swapped.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api, workbook.sourceFormat, workbook.snapshot.name]);

  const handleNew = () => workbook.replaceWorkbook(emptyWorkbook(), null);
  const handleOpen = async () => {
    try {
      const file = await pickXlsxFile();
      if (!file) return;
      await loadSpreadsheetFile(file, api, workbook.replaceWorkbook);
    } catch (err) {
      // Surface failures from the parser / replace flow so they aren't
      // swallowed by React's unhandled-rejection silence on event handlers.
      console.error('[open] failed', err);
      window.alert(`Could not open this file: ${(err as Error)?.message ?? String(err)}`);
    }
  };

  // Save writes back in whatever format the file was opened from, falling
  // back to xlsx for a fresh / empty workbook. Mirrors Excel & LibreOffice.
  const handleSave = async () => {
    if (!api) return;
    const name = workbook.snapshot.name || 'workbook';
    switch (workbook.sourceFormat) {
      case 'ods':
        await saveAsOds(api, name);
        return;
      case 'csv':
        await saveAsCsv(api, name);
        return;
      case 'tsv':
        await saveAsTsv(api, name);
        return;
      case 'xlsx':
      default:
        await saveAsXlsx(api, name);
    }
  };

  const handleExportXlsx = async () => api && saveAsXlsx(api, workbook.snapshot.name || 'workbook');
  const handleExportOds = async () => api && saveAsOds(api, workbook.snapshot.name || 'workbook');
  const handleExportCsv = async () => api && saveAsCsv(api, workbook.snapshot.name || 'workbook');
  const handleExportTsv = async () => api && saveAsTsv(api, workbook.snapshot.name || 'workbook');

  const menus: Record<MenuId, { label: string; items: MenuItem[] }> = {
    file: {
      label: 'File',
      items: [
        { kind: 'item', id: 'new', label: 'New', icon: 'add', shortcut: 'Ctrl+N', onClick: handleNew },
        { kind: 'item', id: 'open', label: 'Open', icon: 'folder_open', shortcut: 'Ctrl+O', onClick: handleOpen },
        { kind: 'item', id: 'save', label: 'Save', icon: 'save', shortcut: 'Ctrl+S', onClick: handleSave },
        {
          kind: 'submenu',
          id: 'export',
          label: 'Export',
          icon: 'ios_share',
          items: [
            { kind: 'item', id: 'export-xlsx', label: '.xlsx', icon: 'description', onClick: handleExportXlsx },
            { kind: 'item', id: 'export-ods', label: '.ods', icon: 'description', onClick: handleExportOds },
            { kind: 'item', id: 'export-csv', label: '.csv', icon: 'description', onClick: handleExportCsv },
            { kind: 'item', id: 'export-tsv', label: '.tsv', icon: 'description', onClick: handleExportTsv },
          ],
        },
        { kind: 'separator', id: 'sep-1' },
        { kind: 'item', id: 'print', label: 'Print', icon: 'print', shortcut: 'Ctrl+P', onClick: () => api && printActiveSheet(api) },
        { kind: 'separator', id: 'sep-2' },
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
        { kind: 'item', id: 'show-all-rows', label: 'Show all rows', icon: 'unfold_more', run: showAllRows },
        { kind: 'separator', id: 'sep-2' },
        { kind: 'item', id: 'tables-panel', label: ui.tablesPanelVisible ? 'Hide Tables panel' : 'Tables panel', icon: 'table_rows', onClick: ui.toggleTablesPanel },
        { kind: 'item', id: 'comments-panel', label: 'Comments panel', icon: 'forum', run: toggleCommentPanel },
      ],
    },
    help: {
      label: 'Help',
      items: [
        {
          kind: 'item',
          id: 'report-bug',
          label: 'Report a bug…',
          icon: 'bug_report',
          onClick: openBugReport,
        },
        { kind: 'separator', id: 'sep-1' },
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
  const [openSubmenu, setOpenSubmenu] = useState<string | null>(null);
  return (
    <>
      {items.map((item) => {
        if (item.kind === 'separator') {
          return <div key={item.id} className="menu__divider" />;
        }
        if (item.kind === 'submenu') {
          const isOpen = openSubmenu === item.id;
          return (
            <div
              key={item.id}
              className="menu__submenu"
              onMouseEnter={() => setOpenSubmenu(item.id)}
              onMouseLeave={() => setOpenSubmenu(null)}
            >
              <button
                type="button"
                role="menuitem"
                className="menu__item menu__item--has-submenu"
                data-testid={`menu-item-${item.id}`}
                aria-haspopup="menu"
                aria-expanded={isOpen}
                onClick={() => setOpenSubmenu(isOpen ? null : item.id)}
              >
                {item.icon && <Icon name={item.icon} size="sm" className="menu__item-icon" />}
                <span>{item.label}</span>
                <Icon name="chevron_right" size="sm" className="menu__item-chevron" />
              </button>
              {isOpen && (
                <div
                  className="menu menu--sub"
                  role="menu"
                  data-testid={`menu-item-${item.id}-popup`}
                >
                  <MenuList items={item.items} onItemClick={onItemClick} />
                </div>
              )}
            </div>
          );
        }
        return (
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
        );
      })}
    </>
  );
}
