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
import { loadPrintOptions, printActiveSheet, savePrintOptions } from './print';
import { PageSetupDialog } from './PageSetupDialog';
import { InsertCellsDialog } from './InsertCellsDialog';
import { openBugReport } from './report-bug';
import { useCollab } from '../collab/collab-context';
import { useLoading } from '../loading-context';
import { useCharts } from '../charts/charts-context';
import {
  buildChartModelForRange,
  getActiveSelectionRange,
  rangeToA1,
} from '../charts/insert-chart';
import { InsertChartDialog } from '../charts/InsertChartDialog';
import { nextChartName } from '../charts/naming';
import { usePivots } from '../pivots/pivots-context';
import { InsertPivotDialog } from '../pivots/InsertPivotDialog';
import { applyPivot } from '../pivots/apply';
import { newPivotId } from '../pivots/types';
import { useOutlineActions } from '../outline/use-outline-actions';
import { useOutline } from '../outline/outline-context';
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
  insertCellsAt,
  deleteCellsAt,
  insertColumnLeft,
  insertColumnRight,
  enterCellEditMode,
  selectEntireColumns,
  selectEntireRows,
  type CellsOpDirection,
  insertComment,
  insertCurrentTime,
  insertHyperlink,
  insertImage,
  insertNewSheet,
  insertRowAbove,
  insertRowBelow,
  insertTable,
  insertTodayDate,
  jumpToFirstCell,
  jumpToLastCell,
  switchToNextSheet,
  switchToPreviousSheet,
  openConditionalFormatting,
  openCustomSort,
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
  const outlineActions = useOutlineActions();
  const outline = useOutline();
  const collab = useCollab();
  const loading = useLoading();
  const charts = useCharts();
  const [open, setOpen] = useState<MenuId | null>(null);
  const [showProperties, setShowProperties] = useState(false);
  const [showAbout, setShowAbout] = useState(false);
  const [showPageSetup, setShowPageSetup] = useState(false);
  const [showInsertChart, setShowInsertChart] = useState(false);
  const [insertChartDefault, setInsertChartDefault] = useState('A1');
  const [showInsertPivot, setShowInsertPivot] = useState(false);
  const [insertPivotDefault, setInsertPivotDefault] = useState('A1');
  const pivots = usePivots();

  // Toolbar's Insert > Chart / Pivot buttons live in a sibling component;
  // they can't reach this dialog state directly, so we dispatch DOM
  // CustomEvents from the toolbar and open the dialogs here. Same pattern
  // keeps both surfaces in sync without lifting state.
  useEffect(() => {
    const openChart = () => {
      if (!api) return;
      const sel = getActiveSelectionRange(api);
      setInsertChartDefault(sel ? rangeToA1(sel) : 'A1');
      setShowInsertChart(true);
    };
    const openPivot = () => {
      if (!api) return;
      const sel = getActiveSelectionRange(api);
      setInsertPivotDefault(sel ? rangeToA1(sel) : 'A1:C10');
      setShowInsertPivot(true);
    };
    document.addEventListener('casual-open-insert-chart', openChart);
    document.addEventListener('casual-open-insert-pivot', openPivot);
    return () => {
      document.removeEventListener('casual-open-insert-chart', openChart);
      document.removeEventListener('casual-open-insert-pivot', openPivot);
    };
  }, [api]);
  // Ctrl++ / Ctrl+- → Excel's Insert / Delete chooser modals. `null`
  // when closed; `'insert'` / `'delete'` when open.
  const [cellsOp, setCellsOp] = useState<'insert' | 'delete' | null>(null);

  const onClose = () => setOpen(null);

  // Intercept Ctrl/Cmd+P globally — the default would print the whole web
  // page (chrome + grid). We print only the active sheet via an offscreen
  // iframe instead. Capture-phase so we beat browser shortcuts on focused
  // inputs as well. Same hook owns Ctrl/Cmd+S → Save (in source format)
  // and Ctrl/Cmd+F → Find & Replace (the latter would otherwise no-op
  // until Univer's find-replace plugin finishes its idle load — see
  // `openFindReplace`, which awaits `ensurePluginByName('findReplace')`).
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const mod = e.metaKey || e.ctrlKey;
      const k = e.key.toLowerCase();
      const inTextInput = (() => {
        const tag = (e.target as HTMLElement | null)?.tagName;
        return tag === 'INPUT' || tag === 'TEXTAREA';
      })();
      if (mod && !e.altKey) {
        // ── File / global ───────────────────────────────────────────
        if (k === 'p' && !e.shiftKey) {
          e.preventDefault();
          if (api) setShowPageSetup(true);
        } else if (k === 's' && !e.shiftKey) {
          e.preventDefault();
          void handleSave();
        } else if (k === 'n' && !e.shiftKey) {
          // Ctrl+N — new workbook. Browser default would open a new
          // window which is almost never what an Excel user wants.
          e.preventDefault();
          handleNew();
        } else if (k === 'o' && !e.shiftKey) {
          // Ctrl+O — open. Browser default is a no-op for users with
          // no app handler; we replace it with the file picker.
          e.preventDefault();
          void handleOpen();
        } else if (k === 'f' && !e.shiftKey) {
          // Skip when focus is in a plain text input — browsers expect
          // Ctrl+F to do in-page find there. The find dialog is for the
          // sheet, not the formula bar / name box.
          if (inTextInput) return;
          e.preventDefault();
          if (api) void openFindReplace(api);
        } else if (k === 'h' && !e.shiftKey) {
          // Ctrl+H — Find & Replace (Excel opens directly on Replace
          // tab; Univer's find-replace plugin doesn't expose a tab
          // switch param, so we open the dialog and let the user
          // click Replace — still better than no binding).
          if (inTextInput) return;
          e.preventDefault();
          if (api) void openFindReplace(api);
        } else if (k === 'k' && !e.shiftKey) {
          // Ctrl+K — insert hyperlink.
          if (inTextInput) return;
          e.preventDefault();
          if (api) insertHyperlink(api);
        }
        // ── Sheet navigation ───────────────────────────────────────
        // Use e.key for PageUp/PageDown since they're not letters
        // and don't normalize via toLowerCase the same way.
        else if (e.key === 'PageUp' && !e.shiftKey) {
          e.preventDefault();
          if (api) switchToPreviousSheet(api);
        } else if (e.key === 'PageDown' && !e.shiftKey) {
          e.preventDefault();
          if (api) switchToNextSheet(api);
        }
        // ── Jump ───────────────────────────────────────────────────
        else if (e.key === 'Home' && !e.shiftKey) {
          // Ctrl+Home — jump to A1.
          if (inTextInput) return;
          e.preventDefault();
          if (api) jumpToFirstCell(api);
        } else if (e.key === 'End' && !e.shiftKey) {
          // Ctrl+End — jump to the bottom-right of the used range.
          if (inTextInput) return;
          e.preventDefault();
          if (api) jumpToLastCell(api);
        }
        // ── Date / time ───────────────────────────────────────────
        else if (e.key === ';' && !e.shiftKey) {
          // Ctrl+; — today's date.
          if (inTextInput) return;
          e.preventDefault();
          if (api) insertTodayDate(api);
        } else if ((e.key === ':' || (e.shiftKey && e.key === ';')) && e.shiftKey) {
          // Ctrl+Shift+: — current time. The key event reports `:`
          // on some keyboards and `;` + shiftKey on others, so we
          // accept both.
          if (inTextInput) return;
          e.preventDefault();
          if (api) insertCurrentTime(api);
        }
      }
      // ── Function keys (no modifier required) ─────────────────────
      if (e.key === 'F11' && e.shiftKey) {
        // Shift+F11 — insert new sheet. Browser may eat F11 alone
        // (full-screen toggle), so the Shift variant is the safe pick.
        if (inTextInput) return;
        e.preventDefault();
        if (api) insertNewSheet(api);
      }
      if (e.key === 'F2' && !mod && !e.shiftKey && !e.altKey) {
        // F2 — drop the active cell into edit mode without clearing
        // its contents. The canonical Excel "edit in place" shortcut.
        if (inTextInput) return;
        e.preventDefault();
        if (api) enterCellEditMode(api);
      }
      // ── Selection: Ctrl+Space, Shift+Space ───────────────────────
      // Both are bare-modifier shortcuts (Ctrl xor Shift, never both).
      // Skip while a text input has focus — Ctrl+Space is also the
      // common autocomplete trigger.
      if (mod && !e.shiftKey && !e.altKey && e.code === 'Space') {
        // Ctrl+Space — select the entire column(s) of the current
        // selection. Excel's most-used "select column" gesture.
        if (inTextInput) return;
        e.preventDefault();
        if (api) selectEntireColumns(api);
      } else if (!mod && e.shiftKey && !e.altKey && e.code === 'Space') {
        // Shift+Space — select the entire row(s).
        if (inTextInput) return;
        e.preventDefault();
        if (api) selectEntireRows(api);
      }
      // ── Insert / Delete cells: Ctrl++ and Ctrl+- ─────────────────
      // The `+` key reports differently across layouts:
      //   - US/UK: e.key === '=' with e.shiftKey === true → Ctrl++
      //   - Numpad +: e.key === '+', e.shiftKey === false
      // Accept both. Same for `-` (key '-' or numpad '-').
      const isPlus = e.code === 'NumpadAdd' || (e.key === '=' && e.shiftKey) || e.key === '+';
      const isMinus = e.code === 'NumpadSubtract' || e.key === '-';
      if (mod && !e.altKey && isPlus) {
        if (inTextInput) return;
        e.preventDefault();
        setCellsOp('insert');
      } else if (mod && !e.altKey && !e.shiftKey && isMinus) {
        if (inTextInput) return;
        e.preventDefault();
        setCellsOp('delete');
      }
    };
    window.addEventListener('keydown', onKey, { capture: true });
    return () => window.removeEventListener('keydown', onKey, { capture: true });
    // handleSave / handleNew / handleOpen read workbook.meta + api
    // from context; api in deps re-binds the handler when the
    // workbook is swapped.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api, workbook.meta]);

  const handleNew = () => workbook.replaceWorkbook(emptyWorkbook(), null);
  const handleOpen = async () => {
    let openedFile: File | null = null;
    try {
      const file = await pickXlsxFile();
      if (!file) return;
      openedFile = file;
      // Open the loading overlay before we await anything heavy — we
      // want the user to see "Reading file" before the parser worker
      // even spins up.
      loading.set({ fileName: file.name, sizeBytes: file.size, phase: 'reading' });
      await loadSpreadsheetFile(file, api, workbook.replaceWorkbook, (phase) =>
        loading.set({ phase }),
      );
      // Give Univer's mount one frame to settle before dropping the
      // overlay — otherwise a fast open shows a blink of the empty
      // grid before the new unit paints.
      requestAnimationFrame(() => loading.set(null));
    } catch (err) {
      // Surface failures in the overlay rather than window.alert so the
      // user can read + copy the actual error message instead of losing
      // it to an OS dialog. `onRetry` reopens the file picker so the
      // user can pick the same (or a different) file without leaving
      // the error card.
      const msg = err instanceof Error ? `${err.name}: ${err.message}` : String(err);
      console.error('[open] failed', err);
      loading.set({
        fileName: openedFile?.name ?? 'workbook',
        sizeBytes: openedFile?.size,
        phase: 'reading',
        error: msg,
        onRetry: () => void handleOpen(),
      });
    }
  };

  // Save writes back in whatever format the file was opened from, falling
  // back to xlsx for a fresh / empty workbook. Mirrors Excel & LibreOffice.
  const handleSave = async () => {
    if (!api) return;
    const name = workbook.meta.name || 'workbook';
    switch (workbook.meta.sourceFormat) {
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
        await saveAsXlsx(api, name, {
          outline: outline.state,
          charts: charts.charts,
          pivots: pivots.pivots,
        });
    }
  };

  const handleExportXlsx = async () =>
    api &&
    saveAsXlsx(api, workbook.meta.name || 'workbook', {
      outline: outline.state,
      charts: charts.charts,
      pivots: pivots.pivots,
    });
  const handleExportOds = async () => api && saveAsOds(api, workbook.meta.name || 'workbook');
  const handleExportCsv = async () => api && saveAsCsv(api, workbook.meta.name || 'workbook');
  const handleExportTsv = async () => api && saveAsTsv(api, workbook.meta.name || 'workbook');

  // Menu structure designed against Office 2024's ribbon + File menu.
  // Every item with a global keyboard binding shows its shortcut on the
  // right of the row; items without one are left bare. Sub-menus are
  // avoided where the items fit in the parent (the previous
  // File → Export submenu pushed common saves behind a hover step).
  const menus: Record<MenuId, { label: string; items: MenuItem[] }> = {
    file: {
      label: 'File',
      items: [
        { kind: 'item', id: 'new', label: 'New', icon: 'add', shortcut: 'Ctrl+N', onClick: handleNew },
        { kind: 'item', id: 'open', label: 'Open…', icon: 'folder_open', shortcut: 'Ctrl+O', onClick: handleOpen },
        { kind: 'separator', id: 'sep-save' },
        { kind: 'item', id: 'save', label: 'Save', icon: 'save', shortcut: 'Ctrl+S', onClick: handleSave },
        // Save As → submenu for the format picker (xlsx default → ods → csv → tsv).
        {
          kind: 'submenu',
          id: 'save-as',
          label: 'Save as',
          icon: 'ios_share',
          items: [
            { kind: 'item', id: 'save-as-xlsx', label: '.xlsx (Excel)',           icon: 'description', onClick: handleExportXlsx },
            { kind: 'item', id: 'save-as-ods',  label: '.ods (OpenDocument)',     icon: 'description', onClick: handleExportOds },
            { kind: 'item', id: 'save-as-csv',  label: '.csv (comma-separated)',  icon: 'description', onClick: handleExportCsv },
            { kind: 'item', id: 'save-as-tsv',  label: '.tsv (tab-separated)',    icon: 'description', onClick: handleExportTsv },
          ],
        },
        { kind: 'separator', id: 'sep-print' },
        { kind: 'item', id: 'print', label: 'Print…', icon: 'print', shortcut: 'Ctrl+P', onClick: () => setShowPageSetup(true) },
        { kind: 'separator', id: 'sep-coedit' },
        ...(collab.roomId
          ? ([
              {
                kind: 'item',
                id: 'download-room',
                label: 'Download a copy (.xlsx)',
                icon: 'download',
                onClick: handleExportXlsx,
              },
              {
                kind: 'item',
                id: 'leave-room',
                label: 'Leave room',
                icon: 'logout',
                onClick: () => {
                  window.location.href = window.location.origin + '/';
                },
              },
            ] as MenuItem[])
          : ([
              {
                kind: 'item',
                id: 'start-room',
                label: 'Share for co-editing…',
                icon: 'group_add',
                onClick: () => ui.openShareRoom(),
              },
            ] as MenuItem[])),
        { kind: 'separator', id: 'sep-props' },
        { kind: 'item', id: 'properties', label: 'Properties…', icon: 'info', onClick: () => setShowProperties(true) },
        { kind: 'item', id: 'about', label: 'About casual sheets', icon: 'help_outline', onClick: () => setShowAbout(true) },
      ],
    },
    edit: {
      label: 'Edit',
      items: [
        { kind: 'item', id: 'undo', label: 'Undo', icon: 'undo', shortcut: 'Ctrl+Z', run: undo },
        { kind: 'item', id: 'redo', label: 'Redo', icon: 'redo', shortcut: 'Ctrl+Y', run: redo },
        { kind: 'separator', id: 'sep-clip' },
        { kind: 'item', id: 'cut', label: 'Cut', icon: 'content_cut', shortcut: 'Ctrl+X', run: actCut },
        { kind: 'item', id: 'copy', label: 'Copy', icon: 'content_copy', shortcut: 'Ctrl+C', run: actCopy },
        { kind: 'item', id: 'paste', label: 'Paste', icon: 'content_paste', shortcut: 'Ctrl+V', run: actPaste },
        { kind: 'separator', id: 'sep-find' },
        { kind: 'item', id: 'find-replace', label: 'Find & Replace…', icon: 'search', shortcut: 'Ctrl+F', run: openFindReplace },
        { kind: 'separator', id: 'sep-cells' },
        // The Insert / Delete dialogs were keyboard-only via Polish #1;
        // surface them in the menu so they're discoverable.
        { kind: 'item', id: 'edit-insert-cells', label: 'Insert cells…', icon: 'add_box', shortcut: 'Ctrl++', onClick: () => setCellsOp('insert') },
        { kind: 'item', id: 'edit-delete-cells', label: 'Delete cells…', icon: 'indeterminate_check_box', shortcut: 'Ctrl+-', onClick: () => setCellsOp('delete') },
        { kind: 'separator', id: 'sep-sel' },
        { kind: 'item', id: 'edit-select-col', label: 'Select column', icon: 'view_column', shortcut: 'Ctrl+Space', onClick: () => api && selectEntireColumns(api) },
        { kind: 'item', id: 'edit-select-row', label: 'Select row', icon: 'view_stream', shortcut: 'Shift+Space', onClick: () => api && selectEntireRows(api) },
        { kind: 'item', id: 'edit-edit-cell', label: 'Edit cell', icon: 'edit', shortcut: 'F2', onClick: () => api && enterCellEditMode(api) },
      ],
    },
    view: {
      label: 'View',
      items: [
        { kind: 'item', id: 'toggle-formula-bar', label: ui.formulaBarVisible ? 'Hide formula bar' : 'Show formula bar', icon: 'functions', onClick: ui.toggleFormulaBar },
        { kind: 'item', id: 'toggle-gridlines', label: 'Gridlines', icon: 'grid_on', onClick: () => api && toggleGridlines(api, true) },
        { kind: 'separator', id: 'sep-freeze' },
        { kind: 'item', id: 'freeze-row', label: 'Freeze top row', icon: 'border_horizontal', run: freezeFirstRow },
        { kind: 'item', id: 'freeze-col', label: 'Freeze first column', icon: 'border_vertical', run: freezeFirstColumn },
        { kind: 'item', id: 'freeze-selection', label: 'Freeze panes (at selection)', icon: 'grid_4x4', run: freezeAtSelection },
        { kind: 'item', id: 'unfreeze', label: 'Unfreeze', icon: 'grid_off', run: unfreezePanes },
        { kind: 'separator', id: 'sep-nav' },
        { kind: 'item', id: 'jump-home', label: 'Jump to A1', icon: 'home', shortcut: 'Ctrl+Home', onClick: () => api && jumpToFirstCell(api) },
        { kind: 'item', id: 'jump-end', label: 'Jump to last cell', icon: 'last_page', shortcut: 'Ctrl+End', onClick: () => api && jumpToLastCell(api) },
        { kind: 'item', id: 'prev-sheet', label: 'Previous sheet', icon: 'navigate_before', shortcut: 'Ctrl+PageUp', onClick: () => api && switchToPreviousSheet(api) },
        { kind: 'item', id: 'next-sheet', label: 'Next sheet', icon: 'navigate_next', shortcut: 'Ctrl+PageDown', onClick: () => api && switchToNextSheet(api) },
        { kind: 'separator', id: 'sep-panels' },
        { kind: 'item', id: 'tables-panel',  label: ui.tablesPanelVisible  ? 'Hide Tables panel'  : 'Tables panel',  icon: 'table_rows', onClick: ui.toggleTablesPanel },
        { kind: 'item', id: 'outline-panel', label: ui.outlinePanelVisible ? 'Hide Outline panel' : 'Outline panel', icon: 'list',       onClick: ui.toggleOutlinePanel },
        { kind: 'item', id: 'charts-panel',  label: ui.chartsPanelVisible  ? 'Hide Charts panel'  : 'Charts panel',  icon: 'bar_chart',  onClick: ui.toggleChartsPanel },
        { kind: 'item', id: 'comments-panel', label: 'Comments panel', icon: 'forum', run: toggleCommentPanel },
      ],
    },
    insert: {
      label: 'Insert',
      items: [
        // High-leverage objects first — what an Excel user reaches for.
        { kind: 'item', id: 'new-sheet', label: 'New sheet', icon: 'add_box', shortcut: 'Shift+F11', run: insertNewSheet },
        { kind: 'item', id: 'insert-table', label: 'Table', icon: 'table_rows', run: insertTable },
        {
          kind: 'item',
          id: 'insert-chart',
          label: 'Chart…',
          icon: 'bar_chart',
          onClick: () => {
            if (!api) return;
            const sel = getActiveSelectionRange(api);
            setInsertChartDefault(sel ? rangeToA1(sel) : 'A1');
            setShowInsertChart(true);
          },
        },
        {
          kind: 'item',
          id: 'insert-pivot',
          label: 'PivotTable…',
          icon: 'pivot_table_chart',
          // Open a configuration dialog (source range + target cell +
          // row field + value field + aggregation), compute the pivot,
          // and write the result as cells at the target location.
          onClick: () => {
            if (!api) return;
            const sel = getActiveSelectionRange(api);
            setInsertPivotDefault(sel ? rangeToA1(sel) : 'A1:C10');
            setShowInsertPivot(true);
          },
        },
        { kind: 'separator', id: 'sep-objects' },
        { kind: 'item', id: 'insert-image', label: 'Image…', icon: 'image', run: insertImage },
        { kind: 'item', id: 'insert-link', label: 'Hyperlink…', icon: 'link', shortcut: 'Ctrl+K', run: insertHyperlink },
        { kind: 'item', id: 'insert-comment', label: 'Comment', icon: 'comment', shortcut: 'Shift+F2', run: insertComment },
        { kind: 'separator', id: 'sep-rowcol' },
        {
          kind: 'submenu',
          id: 'insert-rowcol',
          label: 'Rows & columns',
          icon: 'grid_on',
          items: [
            { kind: 'item', id: 'insert-row-above', label: 'Row above',     icon: 'vertical_align_top',    run: insertRowAbove },
            { kind: 'item', id: 'insert-row-below', label: 'Row below',     icon: 'vertical_align_bottom', run: insertRowBelow },
            { kind: 'item', id: 'insert-col-left',  label: 'Column left',   icon: 'keyboard_tab_rtl',      run: insertColumnLeft },
            { kind: 'item', id: 'insert-col-right', label: 'Column right',  icon: 'keyboard_tab',          run: insertColumnRight },
          ],
        },
        { kind: 'separator', id: 'sep-autofit' },
        { kind: 'item', id: 'autofit-col', label: 'Auto-fit column width', icon: 'settings_ethernet', run: autoFitColumns },
        { kind: 'item', id: 'autofit-row', label: 'Auto-fit row height', icon: 'height', run: autoFitRows },
        { kind: 'separator', id: 'sep-date' },
        { kind: 'item', id: 'insert-today', label: "Today's date", icon: 'today', shortcut: 'Ctrl+;', run: insertTodayDate },
        { kind: 'item', id: 'insert-time', label: 'Current time', icon: 'schedule', shortcut: 'Ctrl+Shift+:', run: insertCurrentTime },
      ],
    },
    format: {
      label: 'Format',
      items: [
        // Number formats live behind a submenu so they don't push the
        // other Format actions off the bottom of the dropdown. The user
        // hovers "Number format" and gets all 10 variants in one place.
        {
          kind: 'submenu',
          id: 'num-format',
          label: 'Number format',
          icon: 'looks_one',
          items: (
            ['general', 'number', 'integer', 'currency', 'accounting', 'percent', 'date', 'time', 'scientific', 'text'] as NumberFormatKey[]
          ).map<MenuItem>((k) => ({
            kind: 'item',
            id: `num-${k}`,
            label: k[0]!.toUpperCase() + k.slice(1),
            icon: 'looks_one',
            onClick: () => api && setNumberFormatByKey(api, k),
          })),
        },
        { kind: 'item', id: 'decimal-up', label: 'Increase decimals', icon: 'add', run: increaseDecimal },
        { kind: 'item', id: 'decimal-down', label: 'Decrease decimals', icon: 'remove', run: decreaseDecimal },
        { kind: 'separator', id: 'sep-cond' },
        { kind: 'item', id: 'conditional-formatting', label: 'Conditional formatting…', icon: 'palette', run: openConditionalFormatting },
        { kind: 'separator', id: 'sep-visibility' },
        // Hide / Unhide grouped — Excel's Format → Visibility submenu.
        {
          kind: 'submenu',
          id: 'visibility',
          label: 'Visibility',
          icon: 'visibility',
          items: [
            { kind: 'item', id: 'hide-row',   label: 'Hide row',     icon: 'visibility_off', run: hideSelectedRows },
            { kind: 'item', id: 'unhide-row', label: 'Unhide row',   icon: 'visibility',     run: unhideSelectedRows },
            { kind: 'item', id: 'hide-col',   label: 'Hide column',  icon: 'visibility_off', run: hideSelectedColumns },
            { kind: 'item', id: 'unhide-col', label: 'Unhide column', icon: 'visibility',     run: unhideSelectedColumns },
          ],
        },
        { kind: 'separator', id: 'sep-fit' },
        { kind: 'item', id: 'autofit-col', label: 'Auto-fit column width', icon: 'settings_ethernet', run: autoFitColumns },
        { kind: 'item', id: 'autofit-row', label: 'Auto-fit row height', icon: 'height', run: autoFitRows },
        { kind: 'separator', id: 'sep-delete' },
        { kind: 'item', id: 'delete-row', label: 'Delete row', icon: 'delete_sweep', run: deleteSelectedRow },
        { kind: 'item', id: 'delete-col', label: 'Delete column', icon: 'folder_delete', run: deleteSelectedColumn },
      ],
    },
    data: {
      label: 'Data',
      items: [
        { kind: 'item', id: 'sort-custom', label: 'Sort range…', icon: 'sort', run: openCustomSort },
        { kind: 'item', id: 'data-validation', label: 'Data validation…', icon: 'rule', run: openDataValidation },
        { kind: 'separator', id: 'sep-clean' },
        { kind: 'item', id: 'text-to-columns', label: 'Text to Columns', icon: 'splitscreen', run: splitTextToColumns },
        { kind: 'item', id: 'remove-duplicates', label: 'Remove Duplicates', icon: 'filter_list_off', run: removeDuplicates },
        { kind: 'item', id: 'show-all-rows', label: 'Show all rows', icon: 'unfold_more', run: showAllRows },
        { kind: 'separator', id: 'sep-outline' },
        { kind: 'item', id: 'group-rows', label: 'Group rows', icon: 'unfold_less', onClick: () => { outlineActions.groupRows(); } },
        { kind: 'item', id: 'group-cols', label: 'Group columns', icon: 'view_week', onClick: () => { outlineActions.groupCols(); } },
        { kind: 'item', id: 'ungroup', label: 'Ungroup', icon: 'unfold_more_double', onClick: () => { outlineActions.ungroupSelection(); } },
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

      {showInsertChart && api && (
        <InsertChartDialog
          api={api}
          defaultSourceA1={insertChartDefault}
          onCancel={() => setShowInsertChart(false)}
          onConfirm={({ source, type }) => {
            const model = buildChartModelForRange(api, source, type);
            if (model) {
              const name = nextChartName(charts.charts);
              charts.insert({ ...model, title: name });
            }
            setShowInsertChart(false);
          }}
        />
      )}

      {showInsertPivot && api && (
        <InsertPivotDialog
          api={api}
          defaultSourceA1={insertPivotDefault}
          onCancel={() => setShowInsertPivot(false)}
          onConfirm={({ source, target, rowFieldColumn, valueFieldColumn, aggregation }) => {
            const wb = api.getActiveWorkbook();
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const ws = wb?.getActiveSheet() as any;
            if (!wb || !ws) {
              setShowInsertPivot(false);
              return;
            }
            const sheetId = ws.getSheetId();
            const model = {
              id: newPivotId(),
              sourceSheetId: sheetId,
              source,
              targetSheetId: sheetId,
              target,
              // Column indices in the dialog are relative to the source range's
              // left edge; the model stores absolute column offsets within the
              // range too. The dialog already gives us the in-range index, which
              // matches what compute.ts expects.
              rows: [{ column: rowFieldColumn }],
              cols: [],
              values: [{ column: valueFieldColumn, agg: aggregation }],
              title: `PivotTable ${pivots.pivots.length + 1}`,
            };
            pivots.insert(model);
            applyPivot(api, model);
            setShowInsertPivot(false);
          }}
        />
      )}

      {showPageSetup && (
        <PageSetupDialog
          initial={loadPrintOptions()}
          onCancel={() => setShowPageSetup(false)}
          onPrint={(options) => {
            savePrintOptions(options);
            setShowPageSetup(false);
            if (api) printActiveSheet(api, options);
          }}
        />
      )}

      {cellsOp && (
        <InsertCellsDialog
          mode={cellsOp}
          onCancel={() => setCellsOp(null)}
          onConfirm={(dir: CellsOpDirection) => {
            const op = cellsOp;
            setCellsOp(null);
            if (!api) return;
            void (op === 'insert' ? insertCellsAt(api, dir) : deleteCellsAt(api, dir));
          }}
        />
      )}

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
