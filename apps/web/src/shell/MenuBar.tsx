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

import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { Icon } from './Icon';
import { openExternal } from './openExternal';
import { promptModal } from './modals';
import { activeSheet, rangeFromA1, sheetId as facadeSheetId } from '../univer-facade';
import { PropertiesDialog } from './PropertiesDialog';
import { FormatCellsDialog } from './FormatCellsDialog';
import { AboutDialog } from './AboutDialog';
import { KeyboardShortcutsDialog } from './KeyboardShortcutsDialog';
import { useSaveStatus } from './save-status-context';
import { formatShortcut } from './shortcut-format';
import { CommandSearchDialog, type CommandSearchItem } from './CommandSearchDialog';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { useToast } from './toast/toast-context';
import { useActivity } from './activity-context';
import { saveNamedVersion } from '../version-history/useVersionHistoryCapture';
import { useUI } from '../use-ui';
import { emptyWorkbook } from '../snapshot';
import {
  loadSpreadsheetFile,
  pickXlsxFile,
  saveAsCsv,
  saveAsOds,
  saveAsPsv,
  saveAsTsv,
  saveAsXlsx,
} from './file-actions';
import { loadPrintOptions, printActiveSheet, savePrintOptions } from './print';
import { exportActiveSheetPdf } from './pdf';
import {
  getAppliedWatermark,
  isWatermarkOn,
  loadWatermarkPref,
  type WatermarkConfig,
} from './watermark';
import { WatermarkDialog } from './WatermarkDialog';
import { PageSetupDialog } from './PageSetupDialog';
import { InsertCellsDialog } from './InsertCellsDialog';
import { PasteSpecialDialog } from './PasteSpecialDialog';
import { NameManagerDialog } from './NameManagerDialog';
import { GoalSeekDialog } from './GoalSeekDialog';
import { GoToSpecialDialog } from './GoToSpecialDialog';
import { RemoveDuplicatesDialog } from './RemoveDuplicatesDialog';
import { CustomListSortDialog } from './CustomListSortDialog';
import { ColorSortDialog } from './ColorSortDialog';
import { EvaluateFormulaDialog } from './EvaluateFormulaDialog';
import { TextToColumnsDialog } from './TextToColumnsDialog';
import { SubtotalsDialog } from './SubtotalsDialog';
import { AdvancedFilterDialog } from './AdvancedFilterDialog';
import { ScenarioManagerDialog } from './ScenarioManagerDialog';
import { MacrosDialog } from './MacrosDialog';
import { InsertSparklineDialog } from '../sparklines/InsertSparklineDialog';
import { useSparklines } from '../sparklines/sparklines-context';
import { flashFill } from './flash-fill';
import { selectDependents, selectPrecedents } from './formula-refs';
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
import { applyPivot, refreshPivot } from '../pivots/apply';
import { newPivotId } from '../pivots/types';
import { computeDrillDown, findPivotAtCell, type DrillDownResult } from '../pivots/drill-down';
import { DrillDownDialog } from '../pivots/DrillDownDialog';
import { useOutlineActions } from '../outline/use-outline-actions';
import { useOutline } from '../outline/outline-context';
import {
  adjustFontSize,
  copy as actCopy,
  cut as actCut,
  decreaseDecimal,
  increaseDecimal,
  openFindReplace,
  paste as actPaste,
  pasteFormattingOnly,
  pasteSpecial,
  redo,
  setBorders,
  setNumberFormatByKey,
  applyCellStyle,
  undo,
  type NumberFormatKey,
  type PasteSpecialMode,
} from './home-tab-actions';
import { applyReadOnly } from '@casualoffice/sheets/sheets';
import {
  protectActiveRange,
  clearRangeProtections,
  protectActiveSheet,
  unprotectActiveSheet,
  isActiveSheetProtected,
} from '../sheets/protection';
import {
  startRecording,
  runMacro,
  saveMacro,
  listMacros,
  nextMacroName,
  findMacroByShortcut,
  type MacroStep,
} from '../sheets/macros';
import { isDesktop } from '../desk-bridge-bootstrap';
import {
  applyAutoFunction,
  autoFitColumns,
  autoFitRows,
  copyFromAbove,
  deleteSelectedColumn,
  deleteSelectedRow,
  forceRecalculate,
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
  showAllRows,
  toggleCommentPanel,
  toggleFilter,
  reapplyFilter,
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

type SheetRange = {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
};

function normalizeRange(range: SheetRange): SheetRange {
  return {
    startRow: range.startRow,
    endRow: range.endRow,
    startColumn: range.startColumn,
    endColumn: range.endColumn,
  };
}

function sameRange(a: SheetRange, b: SheetRange): boolean {
  return (
    a.startRow === b.startRow &&
    a.endRow === b.endRow &&
    a.startColumn === b.startColumn &&
    a.endColumn === b.endColumn
  );
}

function primaryFor(range: SheetRange) {
  return {
    actualRow: range.startRow,
    actualColumn: range.startColumn,
    isMerged: false,
    isMergedMainCell: false,
    startRow: range.startRow,
    startColumn: range.startColumn,
    endRow: range.startRow,
    endColumn: range.startColumn,
    rangeType: 0,
  };
}

function getSelectionRanges(api: FUniver): SheetRange[] {
  const sheet = activeSheet(api);
  if (!sheet) return [];
  // getSelection / getActiveRangeList live on the @univerjs/sheets-ui
  // facade extension, not the core facade — cast is scoped to the one
  // expression that needs the multi-selection API.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const list = (sheet as any).getSelection?.()?.getActiveRangeList?.();
  if (Array.isArray(list) && list.length > 0) {
    return list
      .map((range) => range?.getRange?.())
      .filter((range): range is SheetRange => !!range)
      .map(normalizeRange);
  }
  const active = sheet.getActiveRange?.();
  const range = active?.getRange?.();
  return range ? [normalizeRange(range)] : [];
}

function broadcastAddToSelectionMode(active: boolean): void {
  document.body.dataset.addToSelectionMode = active ? 'true' : 'false';
  document.dispatchEvent(
    new CustomEvent('casual-add-to-selection-mode-changed', {
      detail: { active },
    }),
  );
}

function openContextMenuForActiveCell(api: FUniver): void {
  const canvas = document.querySelector(
    '[id^="univer-sheet-main-canvas_"]',
  ) as HTMLCanvasElement | null;
  if (!canvas) return;
  const sheet = activeSheet(api);
  const range = sheet?.getActiveRange();
  if (!sheet || !range) return;
  // getScrollState + getCellRect live on the sheets-ui facade
  // extension (rendering layer), not the core sheets facade. The
  // cast stays scoped to the one expression that needs them.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const sheetUi = sheet as any;

  let sx = 0;
  let sy = 0;
  const scrollState = sheetUi.getScrollState?.() as
    | {
        sheetViewStartRow?: number;
        sheetViewStartColumn?: number;
        offsetX?: number;
        offsetY?: number;
      }
    | undefined;
  if (scrollState) {
    try {
      const r = scrollState.sheetViewStartRow ?? 0;
      const c = scrollState.sheetViewStartColumn ?? 0;
      const topLeft = sheetUi.getRange(r, c).getCellRect();
      if (topLeft) {
        sx = topLeft.left + (scrollState.offsetX ?? 0);
        sy = topLeft.top + (scrollState.offsetY ?? 0);
      }
    } catch {
      /* ignore and fall back to unscrolled coordinates */
    }
  }

  try {
    const rect = sheetUi.getRange(range.getRow(), range.getColumn()).getCellRect();
    if (!rect) return;
    const canvasRect = canvas.getBoundingClientRect();
    const clientX =
      canvasRect.left + (rect.left - sx) + Math.max(6, Math.min(20, rect.right - rect.left - 6));
    const clientY =
      canvasRect.top + (rect.top - sy) + Math.max(6, Math.min(20, rect.bottom - rect.top - 6));

    const baseInit = {
      bubbles: true,
      cancelable: true,
      composed: true,
      button: 2,
      buttons: 2,
      clientX,
      clientY,
    };
    canvas.dispatchEvent(
      new PointerEvent('pointerdown', {
        ...baseInit,
        pointerId: 1,
        pointerType: 'mouse',
        isPrimary: true,
      }),
    );
    canvas.dispatchEvent(new MouseEvent('mousedown', baseInit));
    canvas.dispatchEvent(
      new PointerEvent('pointerup', {
        ...baseInit,
        pointerId: 1,
        pointerType: 'mouse',
        isPrimary: true,
      }),
    );
    canvas.dispatchEvent(new MouseEvent('mouseup', baseInit));
    canvas.dispatchEvent(
      new MouseEvent('contextmenu', {
        bubbles: true,
        cancelable: true,
        composed: true,
        button: 2,
        buttons: 2,
        clientX,
        clientY,
      }),
    );
  } catch {
    /* grid rect not ready */
  }
}

function collectCommandSearchItems(
  items: MenuItem[],
  api: FUniver | null,
  trail: string[] = [],
  out: CommandSearchItem[] = [],
): CommandSearchItem[] {
  for (const item of items) {
    if (item.kind === 'separator') continue;
    if (item.kind === 'submenu') {
      collectCommandSearchItems(item.items, api, [...trail, item.label], out);
      continue;
    }
    if (item.disabled) continue;
    out.push({
      id: item.id,
      label: item.label,
      path: [...trail, item.label].join(' > '),
      shortcut: item.shortcut,
      run: async () => {
        if (item.run && api) await item.run(api);
        if (item.onClick) item.onClick();
      },
    });
  }
  return out;
}

// Display labels for the Paste Special toast confirmation.
// Kept in sync with PasteSpecialDialog's OPTIONS list.
const PASTE_SPECIAL_LABEL: Record<PasteSpecialMode, string> = {
  all: 'All',
  formulas: 'Formulas',
  values: 'Values',
  formats: 'Formats',
  'col-widths': 'Column widths',
  'no-borders': 'All except borders',
};

export function MenuBar() {
  const api = useUniverAPI();
  const workbook = useWorkbook();
  const toast = useToast();
  const activity = useActivity();
  const saveStatus = useSaveStatus();
  const ui = useUI();
  const outlineActions = useOutlineActions();
  const outline = useOutline();
  const collab = useCollab();
  const loading = useLoading();
  const charts = useCharts();
  const [open, setOpen] = useState<MenuId | null>(null);
  const [showProperties, setShowProperties] = useState(false);
  const [showFormatCells, setShowFormatCells] = useState(false);
  const [showAbout, setShowAbout] = useState(false);
  const [showPageSetup, setShowPageSetup] = useState(false);
  const [showInsertChart, setShowInsertChart] = useState(false);
  const [insertChartDefault, setInsertChartDefault] = useState('A1');
  const [showInsertPivot, setShowInsertPivot] = useState(false);
  const [insertPivotDefault, setInsertPivotDefault] = useState('A1');
  const [showCommandSearch, setShowCommandSearch] = useState(false);
  const [showShortcuts, setShowShortcuts] = useState(false);
  const [watermarkOn, setWatermarkOn] = useState(false);
  const [protectedOn, setProtectedOn] = useState(false);
  // Disposer returned by applyReadOnly; calling it lifts protection.
  const unprotectRef = useRef<(() => void) | null>(null);
  const [showWatermark, setShowWatermark] = useState(false);
  const [watermarkConfig, setWatermarkConfig] = useState<WatermarkConfig>(() =>
    loadWatermarkPref(),
  );
  const addToSelectionModeRef = useRef(false);
  const selectionRangesRef = useRef<SheetRange[]>([]);
  const syntheticSelectionRef = useRef(false);
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
    const openPasteSpecial = () => setShowPasteSpecial(true);
    const openInsertCells = () => setCellsOp('insert');
    const openDeleteCells = () => setCellsOp('delete');
    const openFormatCells = () => setShowFormatCells(true);
    document.addEventListener('casual-open-insert-chart', openChart);
    document.addEventListener('casual-open-insert-pivot', openPivot);
    document.addEventListener('casual-open-paste-special', openPasteSpecial);
    document.addEventListener('casual-open-insert-cells', openInsertCells);
    document.addEventListener('casual-open-delete-cells', openDeleteCells);
    document.addEventListener('casual-open-format-cells', openFormatCells);
    return () => {
      document.removeEventListener('casual-open-insert-chart', openChart);
      document.removeEventListener('casual-open-insert-pivot', openPivot);
      document.removeEventListener('casual-open-paste-special', openPasteSpecial);
      document.removeEventListener('casual-open-insert-cells', openInsertCells);
      document.removeEventListener('casual-open-delete-cells', openDeleteCells);
      document.removeEventListener('casual-open-format-cells', openFormatCells);
    };
  }, [api]);

  // Ctrl++ / Ctrl+- → Excel's Insert / Delete chooser modals. `null`
  // when closed; `'insert'` / `'delete'` when open.
  const [cellsOp, setCellsOp] = useState<'insert' | 'delete' | null>(null);
  const [showPasteSpecial, setShowPasteSpecial] = useState(false);
  const [showNameManager, setShowNameManager] = useState(false);
  const [showGoalSeek, setShowGoalSeek] = useState(false);
  const [iterativeCalc, setIterativeCalc] = useState(false);
  const [showScenarioManager, setShowScenarioManager] = useState(false);
  const [showGoToSpecial, setShowGoToSpecial] = useState(false);
  const [showRemoveDuplicates, setShowRemoveDuplicates] = useState(false);
  const [showCustomListSort, setShowCustomListSort] = useState(false);
  const [showColorSort, setShowColorSort] = useState(false);
  const [showEvaluateFormula, setShowEvaluateFormula] = useState(false);
  const [showTextToColumns, setShowTextToColumns] = useState(false);
  const [showSubtotals, setShowSubtotals] = useState(false);
  const [showAdvancedFilter, setShowAdvancedFilter] = useState(false);
  const [showMacros, setShowMacros] = useState(false);
  const [showInsertSparkline, setShowInsertSparkline] = useState(false);
  const sparklinesCtx = useSparklines();
  const [drillDownResult, setDrillDownResult] = useState<DrillDownResult | null>(null);

  const onClose = () => setOpen(null);

  // Keep the keyboard handler's view of mutable callbacks fresh. The
  // useEffect below intentionally captures `api` + `workbook.meta` in
  // its deps (re-binding on workbook swap), but `handleSave`/`handleNew`/
  // `handleOpen` close over context state (charts, outline, pivots) that
  // updates more often than the deps fire. Without this ref, Ctrl+S
  // would serialize a stale view — e.g., an empty charts list right
  // after inserting a chart.
  const handlersRef = useRef({
    save: async () => {},
    new: () => {},
    open: async () => {},
    drillDown: () => {},
  });

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
      // F5 — Excel's Go To. On the web F5 is browser-refresh, so only claim it
      // when focus is on the grid (not a text input), and open Go To Special.
      if (!mod && !e.altKey && !e.shiftKey && e.key === 'F5' && !inTextInput) {
        e.preventDefault();
        setShowGoToSpecial(true);
        return;
      }
      if (mod && !e.altKey) {
        // ── File / global ───────────────────────────────────────────
        if (k === 'p' && !e.shiftKey) {
          e.preventDefault();
          if (api) setShowPageSetup(true);
        } else if (k === 's' && !e.shiftKey) {
          e.preventDefault();
          void handlersRef.current.save();
        } else if (k === 'n' && !e.shiftKey) {
          // Ctrl+N — new workbook. Browser default would open a new
          // window which is almost never what an Excel user wants.
          e.preventDefault();
          handlersRef.current.new();
        } else if (k === 'o' && !e.shiftKey) {
          // Ctrl+O — open. Browser default is a no-op for users with
          // no app handler; we replace it with the file picker.
          e.preventDefault();
          void handlersRef.current.open();
        } else if (k === 'g' && !e.shiftKey) {
          // Ctrl+G — Go To: focus the Name Box and select its contents so the
          // user can type a reference / named range. (Go To *Special* — select
          // all constants/formulas/blanks/… — is on F5 and the Edit menu.)
          e.preventDefault();
          document.dispatchEvent(new CustomEvent('casual-focus-name-box'));
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
        } else if (k === 'v' && e.shiftKey) {
          // Ctrl+Shift+V — Paste formatting only.
          if (inTextInput) return;
          e.preventDefault();
          if (api) pasteFormattingOnly(api);
        } else if (k === 'k' && !e.shiftKey) {
          // Ctrl+K — insert hyperlink.
          if (inTextInput) return;
          e.preventDefault();
          if (api) insertHyperlink(api);
        } else if (k === 'a' && e.shiftKey) {
          // Ctrl+Shift+A — insert current function's argument template.
          if (inTextInput) return;
          e.preventDefault();
          document.dispatchEvent(new CustomEvent('casual-insert-function-args'));
        } else if (e.code === 'Digit1' && !e.shiftKey) {
          // Ctrl+1 — Excel-style Format Cells dialog.
          if (inTextInput) return;
          e.preventDefault();
          setShowFormatCells(true);
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
      } else if (e.key === 'F10' && e.shiftKey && !mod && !e.altKey) {
        // Shift+F10 — keyboard context menu for the active cell/range.
        // Reuse Univer's existing canvas context-menu path by
        // synthesizing a right-click at the active cell's viewport rect.
        if (inTextInput) return;
        e.preventDefault();
        if (api) openContextMenuForActiveCell(api);
      } else if (e.key === 'F3' && e.shiftKey && !mod && !e.altKey) {
        // Shift+F3 — Insert Function dialog.
        if (inTextInput) return;
        e.preventDefault();
        document.dispatchEvent(new CustomEvent('casual-open-insert-function'));
      } else if (e.key === 'F3' && mod && !e.shiftKey && !e.altKey) {
        // Ctrl+F3 — Name Manager (Excel canonical).
        if (inTextInput) return;
        e.preventDefault();
        setShowNameManager(true);
      } else if (e.key === 'F8' && e.shiftKey && !mod && !e.altKey) {
        // Shift+F8 — Excel's sticky "Add to Selection" mode.
        e.preventDefault();
        addToSelectionModeRef.current = !addToSelectionModeRef.current;
        if (api) selectionRangesRef.current = getSelectionRanges(api);
        broadcastAddToSelectionMode(addToSelectionModeRef.current);
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
      // ── Number format: Ctrl+Shift+1..6 ───────────────────────────
      // Use e.code (Digit1..Digit6) instead of e.key — Shift+1 yields
      // `!` on US, different symbols elsewhere. Digit codes are stable
      // across layouts. The six bindings map to Excel's defaults:
      //   1 Number  2 Time  3 Date  4 Currency  5 Percent  6 Scientific
      if (mod && e.shiftKey && !e.altKey && /^Digit[1-6]$/.test(e.code)) {
        if (inTextInput) return;
        e.preventDefault();
        const which = e.code.slice(-1) as '1' | '2' | '3' | '4' | '5' | '6';
        const fmt: NumberFormatKey =
          which === '1'
            ? 'number'
            : which === '2'
              ? 'time'
              : which === '3'
                ? 'date'
                : which === '4'
                  ? 'currency'
                  : which === '5'
                    ? 'percent'
                    : 'scientific';
        if (api) setNumberFormatByKey(api, fmt);
      }
      // ── Hide / Unhide rows + columns: Ctrl+9, Ctrl+0 ─────────────
      // Same Digit-code trick to dodge layout differences. Shift adds
      // the "unhide" variant. Ctrl+0 in browsers resets zoom — capture
      // is important so we beat the default.
      if (mod && !e.altKey && e.code === 'Digit9') {
        if (inTextInput) return;
        e.preventDefault();
        if (api) (e.shiftKey ? unhideSelectedRows : hideSelectedRows)(api);
      } else if (mod && !e.altKey && e.code === 'Digit0') {
        if (inTextInput) return;
        e.preventDefault();
        if (api) (e.shiftKey ? unhideSelectedColumns : hideSelectedColumns)(api);
      }
      // ── Insert Table: Ctrl+T / Ctrl+L ────────────────────────────
      // Excel binds Ctrl+T (and Ctrl+L) to "create table"; we honour both.
      // Browser defaults (Ctrl+T new tab, Ctrl+L URL bar) are overridden by
      // preventDefault where the browser allows it (always in the desktop app).
      if (mod && !e.altKey && !e.shiftKey && (k === 'l' || k === 't')) {
        if (inTextInput) return;
        e.preventDefault();
        if (api) void insertTable(api);
      }
      // ── AutoSum: Alt+= ──────────────────────────────────────────
      // Inserts `=SUM(<selection>)` one cell past the selection
      // (multi-cell) or `=SUM()` in the active cell (single-cell).
      // No Ctrl/Cmd modifier — purely Alt.
      if (!mod && e.altKey && !e.shiftKey && (e.key === '=' || e.code === 'Equal')) {
        if (inTextInput) return;
        e.preventDefault();
        if (api) applyAutoFunction(api, 'SUM');
      }
      // ── Insert Chart: Alt+F1 ────────────────────────────────────
      // Opens the chart dialog pre-filled with the active selection
      // (same path as Insert > Chart from the menu).
      if (!mod && e.altKey && !e.shiftKey && e.key === 'F1') {
        if (inTextInput) return;
        e.preventDefault();
        if (api) {
          const sel = getActiveSelectionRange(api);
          setInsertChartDefault(sel ? rangeToA1(sel) : 'A1');
          setShowInsertChart(true);
        }
      }
      // ── Force recalc: F9 ────────────────────────────────────────
      // Excel's "recalculate now" — re-runs the engine even for cells
      // whose dependencies haven't changed.
      if (!mod && !e.altKey && !e.shiftKey && e.key === 'F9') {
        if (inTextInput) return;
        e.preventDefault();
        if (api) forceRecalculate(api);
      }
      // ── Font size: Ctrl+Shift+> / Ctrl+Shift+< ──────────────────
      // Excel's "grow/shrink font". The Period/Comma codes are layout-
      // stable; e.key on US is `>` / `<` but localizes elsewhere.
      if (mod && e.shiftKey && !e.altKey && e.code === 'Period') {
        if (inTextInput) return;
        e.preventDefault();
        if (api) adjustFontSize(api, +1);
      } else if (mod && e.shiftKey && !e.altKey && e.code === 'Comma') {
        if (inTextInput) return;
        e.preventDefault();
        if (api) adjustFontSize(api, -1);
      }
      // ── AutoFilter toggle: Ctrl+Shift+L ─────────────────────────
      // Same key (`l`) as Insert Table — distinguished by Shift.
      if (mod && e.shiftKey && !e.altKey && k === 'l') {
        if (inTextInput) return;
        e.preventDefault();
        if (api) void toggleFilter(api);
      }
      // ── Re-apply filter: Ctrl+Alt+L ─────────────────────────────
      // Excel's "re-evaluate the active filter" — rows edited to no
      // longer match stay visible until this runs. Distinguished from
      // Ctrl+Shift+L (toggle filter) by the Alt modifier.
      if (mod && e.altKey && !e.shiftKey && k === 'l') {
        if (inTextInput) return;
        e.preventDefault();
        if (api) void reapplyFilter(api);
      }
      // ── Paste Special: Ctrl+Alt+V ───────────────────────────────
      // Excel's Paste Special chooser. Opens the dialog rather than
      // pasting blindly — the mode (values / formulas / formats / …)
      // changes the result enough to warrant a confirm step.
      if (mod && e.altKey && !e.shiftKey && k === 'v') {
        if (inTextInput) return;
        e.preventDefault();
        setShowPasteSpecial(true);
      }
      // ── Flash Fill: Ctrl+E ──────────────────────────────────────
      // Excel's heuristic pattern-fill — reads examples typed in the
      // current column, infers a transform from the column to its
      // left, and fills the blanks. See `flash-fill.ts` for the
      // algorithm + supported patterns. No-op when no pattern can be
      // inferred (we don't beep at the user; the silent no-op matches
      // Excel's behaviour for unclear cases).
      if (mod && !e.altKey && !e.shiftKey && k === 'e') {
        if (inTextInput) return;
        e.preventDefault();
        if (api) runFlashFillWithToast(api);
      }
      // ── Pivot drill-down: Ctrl+Shift+D ─────────────────────────
      // Pops the source rows that contributed to the selected pivot
      // cell. Silent no-op when the selection isn't inside a pivot's
      // recorded output extent (matches Excel's "Show Details" silent
      // miss on non-pivot cells). Routed through `handlersRef` so
      // a freshly inserted pivot is visible to the shortcut without
      // re-binding the keydown listener.
      if (mod && e.shiftKey && !e.altKey && k === 'd') {
        if (inTextInput) return;
        e.preventDefault();
        handlersRef.current.drillDown();
      }
      // ── Show Formulas toggle: Ctrl+` ────────────────────────────
      // Excel's `Ctrl+grave-accent` flip. The `e.key` for that key is
      // a literal backtick on US/UK layouts; `e.code` is `Backquote`
      // for layout-stable matching.
      if (mod && !e.altKey && !e.shiftKey && (e.key === '`' || e.code === 'Backquote')) {
        if (inTextInput) return;
        e.preventDefault();
        ui.toggleShowFormulas();
      }
      // ── Precedent / dependent navigation: Ctrl+[ / Ctrl+] ──────
      // Excel's "trace precedents / dependents" via selection. The
      // brackets show up as `[` / `]` in the key field. See
      // `formula-refs.ts` for the regex-based ref extractor.
      if (mod && !e.altKey && !e.shiftKey && e.key === '[') {
        if (inTextInput) return;
        e.preventDefault();
        if (api) selectPrecedents(api);
      } else if (mod && !e.altKey && !e.shiftKey && e.key === ']') {
        if (inTextInput) return;
        e.preventDefault();
        if (api) selectDependents(api);
      }
      // ── Outline border: Ctrl+Shift+& (US) / Ctrl+Shift+7 ────────
      // Both map to Excel's "outside border". US keyboards report
      // Shift+7 as `&`; Digit7 is the layout-stable signal.
      if (mod && e.shiftKey && !e.altKey && e.code === 'Digit7') {
        if (inTextInput) return;
        e.preventDefault();
        if (api) setBorders(api, 'outside');
      }
      // ── Copy from above: Ctrl+' / Ctrl+Shift+' ──────────────────
      // Excel's "fill down from row above" pair. Quote (`'`) is layout-
      // stable on US — falling back to e.key for non-US.
      if (mod && !e.altKey && (e.key === "'" || e.code === 'Quote')) {
        if (inTextInput) return;
        e.preventDefault();
        if (api) copyFromAbove(api, e.shiftKey ? 'value' : 'formula');
      }
      // ── Save As xlsx: Alt+F2 ────────────────────────────────────
      if (!mod && e.altKey && !e.shiftKey && e.key === 'F2') {
        if (inTextInput) return;
        e.preventDefault();
        void handleExportXlsx();
      }
      // ── Tell Me / command search: Alt+Q ────────────────────────
      if (!mod && e.altKey && !e.shiftKey && k === 'q') {
        e.preventDefault();
        setShowCommandSearch(true);
      }
      // ── Command palette: Ctrl+Shift+P (VS Code / Linear convention)
      // UX_AUDIT.md §4.2 / Phase 4 #15. Adds modern muscle memory
      // without touching Excel's Ctrl+K = Insert Link — both chords
      // land on the same `CommandSearchDialog` so there's a single
      // surface, not two competing ones.
      if (mod && !e.altKey && e.shiftKey && k === 'p') {
        e.preventDefault();
        setShowCommandSearch(true);
      }
      // ── Run a macro bound to Ctrl+Shift+<letter> ─────────────────
      // Excel's macro shortcut. Letters L/D/P are app-reserved and can't
      // be assigned (see RESERVED_MACRO_LETTERS), so this branch sitting
      // after them never double-fires — it only acts when a macro holds
      // the pressed letter. e.code (KeyA…KeyZ) dodges layout shifts.
      if (mod && e.shiftKey && !e.altKey && /^Key[A-Z]$/.test(e.code)) {
        if (!inTextInput) {
          const macro = findMacroByShortcut(e.code.slice(-1));
          if (macro) {
            e.preventDefault();
            void handleRunMacro(macro.name);
          }
        }
      }
      // ── Keyboard shortcuts cheat sheet: Ctrl+/ ─────────────────
      // Mirrors Excel for the Web. `?` (Shift+/) is the Google Docs
      // standard but conflicts with the formula-bar editor where the
      // user might type a literal `?`, so we only honour it outside
      // a text-input context — `Ctrl+/` always works.
      if (mod && !e.altKey && !e.shiftKey && (e.key === '/' || e.code === 'Slash')) {
        e.preventDefault();
        setShowShortcuts(true);
      }
      if (!mod && !e.altKey && e.shiftKey && (e.key === '?' || e.key === '/')) {
        if (inTextInput) return;
        e.preventDefault();
        setShowShortcuts(true);
      }
      // ── Close workbook / leave room: Ctrl+W ─────────────────────
      // Browser default is "close tab" — preventDefault and route to /
      // instead so a co-edit room can be left without killing the tab.
      // No-op meaning preserved single-user: just resets the workbook.
      if (mod && !e.altKey && !e.shiftKey && k === 'w') {
        if (inTextInput) return;
        e.preventDefault();
        if (collab.roomId) {
          window.location.href = window.location.origin + '/';
        } else {
          handleNew();
        }
      }
    };
    document.addEventListener('keydown', onKey, { capture: true });
    return () => document.removeEventListener('keydown', onKey, { capture: true });
    // handleSave / handleNew / handleOpen read workbook.meta + api
    // from context; api in deps re-binds the handler when the
    // workbook is swapped.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api, workbook.meta]);

  useEffect(() => {
    if (!api) return;
    selectionRangesRef.current = getSelectionRanges(api);
    const onAddSelectionA1 = (event: Event) => {
      if (!addToSelectionModeRef.current) return;
      const target = (event as CustomEvent<{ target?: string }>).detail?.target?.trim();
      if (!target) return;
      const wb = api.getActiveWorkbook();
      const sheet = wb?.getActiveSheet();
      if (!wb || !sheet) return;
      try {
        const aRange = rangeFromA1(sheet, target);
        if (!aRange) return;
        const nextRange = normalizeRange(aRange.getRange());
        const next = [...selectionRangesRef.current];
        if (!next.some((existing) => sameRange(existing, nextRange))) next.push(nextRange);
        selectionRangesRef.current = next;
        api.executeCommand('sheet.operation.set-selections', {
          unitId: wb.getId(),
          subUnitId: facadeSheetId(sheet),
          selections: next.map((range, index) => ({
            range,
            primary: index === next.length - 1 ? primaryFor(range) : null,
            style: null,
          })),
        });
      } catch {
        /* invalid name-box target */
      }
    };
    const disposable = api.addEvent(api.Event.CommandExecuted, (e) => {
      const id = (e as { id?: string }).id;
      if (id !== 'sheet.operation.set-selections') return;
      const current = getSelectionRanges(api);
      if (syntheticSelectionRef.current) {
        selectionRangesRef.current = current;
        return;
      }
      if (!addToSelectionModeRef.current) {
        selectionRangesRef.current = current;
        return;
      }
      const next = [...selectionRangesRef.current];
      for (const range of current) {
        if (!next.some((existing) => sameRange(existing, range))) next.push(range);
      }
      if (next.length === selectionRangesRef.current.length) {
        selectionRangesRef.current = current;
        return;
      }
      const wb = api.getActiveWorkbook();
      const sheet = activeSheet(api);
      if (!wb || !sheet) return;
      syntheticSelectionRef.current = true;
      queueMicrotask(() => {
        api.executeCommand('sheet.operation.set-selections', {
          unitId: wb.getId(),
          subUnitId: facadeSheetId(sheet),
          selections: next.map((range, index) => ({
            range,
            primary: index === next.length - 1 ? primaryFor(range) : null,
            style: null,
          })),
        });
        queueMicrotask(() => {
          syntheticSelectionRef.current = false;
          selectionRangesRef.current = next;
        });
      });
    });
    document.addEventListener('casual-add-selection-a1', onAddSelectionA1);
    return () => {
      disposable.dispose();
      document.removeEventListener('casual-add-selection-a1', onAddSelectionA1);
    };
  }, [api]);

  // Seed the View → Confidential watermark toggle from the persisted
  // WatermarkService config so the checkmark survives a reload (the plugin
  // re-hydrates the layer from localStorage on boot — see shell/watermark.ts).
  useEffect(() => {
    if (!api) return;
    let cancelled = false;
    void isWatermarkOn(api).then((on) => {
      if (!cancelled) setWatermarkOn(on);
    });
    // Seed the dialog from the currently-applied config when on, else fall
    // back to the user's last-chosen preference from localStorage.
    void getAppliedWatermark(api).then((applied) => {
      if (!cancelled && applied) setWatermarkConfig(applied);
    });
    return () => {
      cancelled = true;
    };
  }, [api]);

  const handleNew = () => {
    workbook.replaceWorkbook(emptyWorkbook(), null);
    // Desktop: a brand-new blank workbook must not stay bound to the path of
    // the file this window previously had open — otherwise the next Save would
    // overwrite that file on disk with the empty workbook. Clear the bound path
    // so save() falls through to saveAs() (prompts for a location), matching the
    // untitled-document save semantics.
    const bridge = typeof window !== 'undefined' ? window.__deskApp__ : undefined;
    if (bridge?.isDesktop) bridge.filePath = null;
  };
  const handleOpen = async () => {
    // Desktop: route File → Open through the shell's native dialog + "this
    // window or a new window?" prompt (mirrors the launcher). The bridge does
    // the open itself — new window, or navigating this one — so the in-window
    // browser-picker flow below is web-only.
    const deskBridge = typeof window !== 'undefined' ? window.__deskApp__ : undefined;
    if (deskBridge?.isDesktop && deskBridge.openViaMenu) {
      await deskBridge.openViaMenu();
      return;
    }
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
      // Desktop: the opened file came from the browser picker, which has no
      // real filesystem path, and it replaced the workbook in-window. Unbind
      // the previously-open file so the next Save can't overwrite it with this
      // content — save() falls through to saveAs() (prompts for a location).
      const bridge = typeof window !== 'undefined' ? window.__deskApp__ : undefined;
      if (bridge?.isDesktop) bridge.filePath = null;
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

  // Pivot drill-down: examine the active selection, find the pivot
  // whose output rectangle contains it, and pop the contributing rows.
  // No-op when the selection isn't inside a known pivot extent.
  const runDrillDown = () => {
    if (!api) return;
    const wb = api.getActiveWorkbook();
    const ws = activeSheet(api);
    const range = ws?.getActiveRange();
    if (!wb || !ws || !range) return;
    const sheetId = facadeSheetId(ws);
    if (!sheetId) return;
    const row = range.getRow();
    const col = range.getColumn();
    const pivot = findPivotAtCell(pivots.pivots, sheetId, row, col);
    if (!pivot) {
      console.debug('[drill-down] no pivot at', { sheetId, row, col });
      return;
    }
    const result = computeDrillDown(api, pivot, row, col);
    if (result) setDrillDownResult(result);
  };

  // Save writes back in whatever format the file was opened from, falling
  // back to xlsx for a fresh / empty workbook. Mirrors Excel & LibreOffice.
  //
  // Wrapped in try/catch + toast so the user gets feedback for both
  // success (the browser download bar can be hidden / off-screen on
  // some setups) and failure (network blip during the worker handoff,
  // out-of-memory on huge files, etc.). Pre-toast, errors were
  // swallowed at the call site and the user saw nothing.
  const handleSave = async () => {
    if (!api) return;
    // Draft skip — UX_AUDIT.md §5. A workbook opened from `/sheet/new`
    // (no server id yet) that the user has never typed in shouldn't
    // materialise a server row on Ctrl+S; that's the silent-junk-row
    // bug that prompted this audit. Once `hasUserEdited` flips, every
    // subsequent Save runs normally — including a "save as empty doc"
    // after they delete every cell, since that delete itself counts.
    if ((workbook.meta.serverFileId ?? null) === null && workbook.meta.hasUserEdited !== true) {
      toast.info('Nothing to save yet — type something first.');
      return;
    }
    saveStatus.markSaving();
    const name = workbook.meta.name || 'workbook';
    // Server-backed save threads the tracked file id + etag in; the
    // FileSource does the in-place PUT and returns the new etag, which
    // we write back into the workbook context so the next Save sees
    // the current version. Browser modes leave these as null and
    // saveAsXlsx falls through to its download / FSA path.
    const serverFileId = workbook.meta.serverFileId ?? null;
    const serverEtag = workbook.meta.serverEtag ?? null;
    const onServerEtag = (etag: string | null) => workbook.updateServerEtag(etag);
    // Bind the server-minted id back on the first create-save so the
    // next save takes the in-place PUT path instead of a duplicating
    // POST. Also rewrites /sheet/new → /sheet/<id> in the URL bar.
    // UX_AUDIT.md §2.3.
    const onServerFileId = (fileId: string) => workbook.updateServerFileId(fileId);
    const onConflict = (expectedEtag: string) => {
      // Conflicts (stale etag) are NOT retryable — blindly re-PUTting
      // would clobber the newer server copy. Keep the bare error + the
      // reload guidance; no Retry button (plain bridge entry via
      // toast.error). The user must reload to pull the latest first.
      toast.error(
        `This file was changed elsewhere (server has ${expectedEtag.slice(0, 8)}…). ` +
          `Reload to pull the latest and try again.`,
      );
    };
    // The actual save, factored out so a failed save can be retried by
    // re-running the SAME call. saveAsXlsx re-snapshots the workbook on
    // every call, so the retry captures fresh data — no stale closure.
    const runSave = () => {
      switch (workbook.meta.sourceFormat) {
        case 'ods':
          return saveAsOds(api, name);
        case 'csv':
          return saveAsCsv(api, name);
        case 'tsv':
          return saveAsTsv(api, name);
        case 'psv':
          return saveAsPsv(api, name);
        case 'xlsx':
        default:
          return saveAsXlsx(api, name, {
            outline: outline.state,
            charts: charts.charts,
            pivots: pivots.pivots,
            sparklines: sparklinesCtx.sparklines,
            serverFileId,
            serverEtag,
            onServerEtag,
            onServerFileId,
            onConflict,
          });
      }
    };
    try {
      await runSave();
      const ext = (workbook.meta.sourceFormat || 'xlsx').toLowerCase();
      const displayName = /\.(xlsx|ods|csv|tsv|psv)$/i.test(name) ? name : `${name}.${ext}`;
      toast.success(`Saved ${displayName}`);
      // Clears `hasUserEdited` so the logout dirty-check stays quiet
      // for users who saved and walked away (UX_AUDIT.md §2.14). The
      // EditTracker flips the flag back on the very next mutation.
      workbook.markSaved();
      // Drive the title-bar SaveStatusPill (UX_AUDIT.md §4.3).
      saveStatus.markSaved();
    } catch (err) {
      // A thrown error is a transient/operational failure (network
      // blip, serializer OOM, disk full) — these ARE retryable. Push a
      // persistent activity entry carrying the retry closure; suppress
      // the toast's own bridge entry so the log shows ONE row (with the
      // Retry button) instead of two. (Conflicts never throw — they go
      // through onConflict above.)
      const msg = err instanceof Error ? err.message : String(err);
      toast.error(`Couldn't save: ${msg}`, { skipActivityLog: true });
      activity.pushErrorWithRetry(`Couldn't save: ${msg}`, runSave, 'save');
      saveStatus.markError(msg);
    }
  };

  // Export = "Save As" to a specific format. Same surface as the
  // primary Save handler — wrap each in try/catch + toast so the
  // user sees confirmation of the download or a real error if the
  // serializer choked. Mirrors the handleSave pattern.
  // Outcome-aware Flash Fill wrapper. The underlying helper returns
  // one of four results; we want to acknowledge the win ("Filled N
  // cells") AND explain the silent-failure cases ("no pattern",
  // "no source column", etc.) so the user isn't left wondering why
  // nothing happened. Used by both the Ctrl+E shortcut and the
  // Data → Flash Fill menu item.
  const runFlashFillWithToast = (univer: FUniver) => {
    const result = flashFill(univer);
    switch (result.status) {
      case 'filled':
        toast.success(
          `Flash Fill: filled ${result.count} ${result.count === 1 ? 'cell' : 'cells'}`,
        );
        return;
      case 'no-pattern':
        toast.info("Flash Fill: couldn't infer a pattern from the examples");
        return;
      case 'no-source':
        toast.info('Flash Fill: needs a column to the left to derive from');
        return;
      case 'no-examples':
        toast.info('Flash Fill: type a few examples in the column first');
        return;
    }
  };

  const exportAs = async (
    format: 'xlsx' | 'ods' | 'csv' | 'tsv' | 'psv',
    runner: () => Promise<unknown>,
  ) => {
    if (!api) return;
    const name = workbook.meta.name || 'workbook';
    const displayName = new RegExp(`\\.${format}$`, 'i').test(name) ? name : `${name}.${format}`;
    try {
      await runner();
      toast.success(`Exported ${displayName}`);
    } catch (err) {
      // Export is local (no etag/conflict path) so every failure is
      // retryable — `runner` re-snapshots on each call. One activity
      // entry (with Retry); suppress the toast's bare bridge entry.
      const msg = err instanceof Error ? err.message : String(err);
      toast.error(`Couldn't export ${displayName}: ${msg}`, { skipActivityLog: true });
      activity.pushErrorWithRetry(
        `Couldn't export ${displayName}: ${msg}`,
        async () => {
          await runner();
        },
        'export',
      );
    }
  };

  const handleExportXlsx = () =>
    exportAs('xlsx', () =>
      saveAsXlsx(api!, workbook.meta.name || 'workbook', {
        outline: outline.state,
        charts: charts.charts,
        pivots: pivots.pivots,
        sparklines: sparklinesCtx.sparklines,
        // Save As / Export → always open the native picker so the user
        // chooses where it lands (never silently overwrite the bound file).
        forcePrompt: true,
      }),
    );

  // Keep the keyboard-listener's handlers in sync with the latest
  // closure after every render. See `handlersRef` declaration above
  // for the rationale.
  handlersRef.current = {
    save: handleSave,
    new: () => {
      handleNew();
    },
    open: handleOpen,
    drillDown: runDrillDown,
  };
  const handleExportOds = () =>
    exportAs('ods', () => saveAsOds(api!, workbook.meta.name || 'workbook'));
  const handleExportCsv = () =>
    exportAs('csv', () => saveAsCsv(api!, workbook.meta.name || 'workbook'));
  const handleExportTsv = () =>
    exportAs('tsv', () => saveAsTsv(api!, workbook.meta.name || 'workbook'));
  const handleExportPsv = () =>
    exportAs('psv', () => saveAsPsv(api!, workbook.meta.name || 'workbook'));

  const handleExportPdf = async () => {
    if (!api) return;
    const r = await exportActiveSheetPdf(api);
    if (r.ok) toast.success('Exported PDF');
    else if (r.reason === 'empty') toast.info('Nothing to export yet — add some data first.');
    // 'cancelled' (user dismissed the native Save dialog) → no toast.
  };

  const toggleProtect = () => {
    if (!api) return;
    if (unprotectRef.current) {
      unprotectRef.current();
      unprotectRef.current = null;
      setProtectedOn(false);
      toast.info('Protection removed — the workbook is editable');
    } else {
      const unitId = api.getActiveWorkbook()?.getId();
      if (!unitId) return;
      // Reuse the SDK's read-only engine (command veto + permission flip).
      unprotectRef.current = applyReadOnly(api, unitId);
      setProtectedOn(true);
      toast.success('Protected — the workbook is read-only');
    }
  };

  // Range protection (T4.4): lock the current selection while the rest of the
  // sheet stays editable — finer than the workbook read-only toggle above.
  const handleProtectRange = async () => {
    if (!api) return;
    const r = await protectActiveRange(api);
    // Collab-protection model: the protector keeps editing; other editors are
    // blocked. Wording makes that intent explicit (vs "locked", which reads as
    // locked-for-everyone).
    if (r.ok) toast.success(`Protected ${r.a1} — other editors can’t change it`);
    else if (r.reason === 'no-selection') toast.info('Select the cells to protect first');
    else if (r.reason === 'overlap')
      toast.info('That selection overlaps an existing protected range');
    else toast.info('Range protection unavailable');
  };

  const handleRemoveRangeProtections = async () => {
    if (!api) return;
    const n = await clearRangeProtections(api);
    if (n > 0) toast.success(`Removed protection from ${n} range${n === 1 ? '' : 's'}`);
    else toast.info('No protected ranges on this sheet');
  };

  // Per-sheet protection (collab model): other editors can't change this sheet;
  // sibling sheets stay editable. Distinct from the workbook read-only toggle.
  const [sheetProtectedOn, setSheetProtectedOn] = useState(false);
  const toggleSheetProtect = async () => {
    if (!api) return;
    if (isActiveSheetProtected(api)) {
      const ok = await unprotectActiveSheet(api);
      if (ok) {
        setSheetProtectedOn(false);
        toast.info('Sheet protection removed');
      }
    } else {
      const ok = await protectActiveSheet(api);
      if (ok) {
        setSheetProtectedOn(true);
        toast.success('Sheet protected — other editors can’t change it');
      } else toast.info('Sheet protection unavailable');
    }
  };

  // Macros (Phase 5): record command-bus mutations → named macro → replay.
  const [macroRecording, setMacroRecording] = useState(false);
  const [macroTick, setMacroTick] = useState(0); // bump to refresh the saved list
  const macroStopRef = useRef<(() => MacroStep[]) | null>(null);
  const handleRecordMacro = () => {
    if (!api) return;
    macroStopRef.current = startRecording(api).stop;
    setMacroRecording(true);
    toast.info('Recording macro — make your edits, then Stop');
  };
  const handleStopMacro = () => {
    const stop = macroStopRef.current;
    macroStopRef.current = null;
    setMacroRecording(false);
    if (!stop) return;
    const steps = stop();
    if (steps.length === 0) {
      toast.info('Nothing recorded — no cell changes captured');
      return;
    }
    const name = nextMacroName();
    saveMacro({ name, steps, createdAt: 0 });
    setMacroTick((t) => t + 1);
    toast.success(`Saved ${name} (${steps.length} step${steps.length === 1 ? '' : 's'})`);
  };
  const handleRunMacro = async (name: string) => {
    if (!api) return;
    const macro = listMacros().find((m) => m.name === name);
    if (!macro) return;
    const n = await runMacro(api, macro.steps);
    toast.success(`Ran ${name} (${n} step${n === 1 ? '' : 's'})`);
  };
  void macroTick; // referenced so the saved-macro submenu rebuilds after a save

  // Menu structure designed against Office 2024's ribbon + File menu.
  // Every item with a global keyboard binding shows its shortcut on the
  // right of the row; items without one are left bare. Sub-menus are
  // avoided where the items fit in the parent (the previous
  // File → Export submenu pushed common saves behind a hover step).
  const menus: Record<MenuId, { label: string; items: MenuItem[] }> = {
    file: {
      label: 'File',
      items: [
        {
          kind: 'item',
          id: 'new',
          label: 'New',
          icon: 'add',
          shortcut: 'Ctrl+N',
          onClick: handleNew,
        },
        {
          kind: 'item',
          id: 'open',
          label: 'Open…',
          icon: 'folder_open',
          shortcut: 'Ctrl+O',
          onClick: handleOpen,
        },
        { kind: 'separator', id: 'sep-save' },
        {
          kind: 'item',
          id: 'save',
          label: 'Save',
          icon: 'save',
          shortcut: 'Ctrl+S',
          onClick: handleSave,
        },
        // Save As → submenu for the format picker (xlsx default → ods → csv → tsv).
        {
          kind: 'submenu',
          id: 'save-as',
          label: 'Save as',
          icon: 'ios_share',
          items: [
            {
              kind: 'item',
              id: 'save-as-xlsx',
              label: '.xlsx (Excel)',
              icon: 'description',
              onClick: handleExportXlsx,
            },
            {
              kind: 'item',
              id: 'save-as-ods',
              label: '.ods (OpenDocument)',
              icon: 'description',
              onClick: handleExportOds,
            },
            {
              kind: 'item',
              id: 'save-as-csv',
              label: '.csv (comma-separated)',
              icon: 'description',
              onClick: handleExportCsv,
            },
            {
              kind: 'item',
              id: 'save-as-tsv',
              label: '.tsv (tab-separated)',
              icon: 'description',
              onClick: handleExportTsv,
            },
            {
              kind: 'item',
              id: 'save-as-psv',
              label: '.psv (pipe-separated)',
              icon: 'description',
              onClick: handleExportPsv,
            },
          ],
        },
        { kind: 'separator', id: 'sep-print' },
        {
          kind: 'item',
          id: 'export-pdf',
          label: 'Download as PDF (.pdf)',
          icon: 'picture_as_pdf',
          onClick: handleExportPdf,
        },
        {
          kind: 'item',
          id: 'print',
          label: 'Print…',
          icon: 'print',
          shortcut: 'Ctrl+P',
          onClick: () => setShowPageSetup(true),
        },
        {
          kind: 'item',
          id: 'set-print-area',
          label: 'Set Print Area to selection',
          icon: 'crop_free',
          onClick: () => {
            if (!api) return;
            const sel = getActiveSelectionRange(api);
            if (!sel) {
              toast.error('Select a range first');
              return;
            }
            const a1 = rangeToA1(sel);
            const prev = loadPrintOptions();
            savePrintOptions({ ...prev, printArea: a1 });
            // Print Area is invisible until you next open Print, so
            // surface it as a confirmation. Offer Undo back to the
            // previous value (null OR the prior range) since it's
            // easy to mis-click on a selection that isn't what you
            // wanted.
            toast.success(`Print Area set to ${a1}`, {
              action: {
                label: 'Undo',
                onClick: () =>
                  savePrintOptions({ ...loadPrintOptions(), printArea: prev.printArea }),
              },
            });
          },
        },
        {
          kind: 'item',
          id: 'clear-print-area',
          label: 'Clear Print Area',
          icon: 'border_clear',
          onClick: () => {
            const prev = loadPrintOptions();
            if (!prev.printArea) {
              toast.info('No Print Area to clear');
              return;
            }
            savePrintOptions({ ...prev, printArea: null });
            toast.success('Print Area cleared', {
              action: {
                label: 'Undo',
                onClick: () =>
                  savePrintOptions({ ...loadPrintOptions(), printArea: prev.printArea }),
              },
            });
          },
        },
        // Co-editing affordances — omitted entirely in the desktop build
        // (single-user, local-file app with no collab server).
        ...(isDesktop()
          ? ([] as MenuItem[])
          : ([
              { kind: 'separator', id: 'sep-coedit' },
              ...(collab.roomId
                ? [
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
                  ]
                : [
                    {
                      kind: 'item',
                      id: 'start-room',
                      label: 'Share for co-editing…',
                      icon: 'group_add',
                      onClick: () => ui.openShareRoom(),
                    },
                  ]),
            ] as MenuItem[])),
        { kind: 'separator', id: 'sep-version' },
        {
          kind: 'item',
          id: 'save-version',
          label: 'Save version…',
          icon: 'bookmark_add',
          onClick: async () => {
            const name = await promptModal({
              title: 'Save version',
              label: 'Name this version (e.g. "Q3 draft", "before pivot redo")',
              defaultValue: 'New version',
              confirmLabel: 'Save',
            });
            if (!name || !api) return;
            const data = api.getActiveWorkbook()?.save() as unknown as
              | Parameters<typeof saveNamedVersion>[0]
              | undefined;
            if (!data) {
              toast.error("Couldn't read workbook state to save a version");
              return;
            }
            const trimmed = name.trim() || 'Untitled version';
            void saveNamedVersion(data, trimmed, workbook.meta.sourceFormat ?? null)
              .then(() => {
                toast.success(`Saved version "${trimmed}"`, {
                  action: {
                    label: 'Open history',
                    onClick: () => {
                      if (!ui.historyPanelVisible) ui.toggleHistoryPanel();
                    },
                  },
                });
              })
              .catch((err: unknown) => {
                const msg = err instanceof Error ? err.message : String(err);
                toast.error(`Couldn't save version: ${msg}`);
              });
          },
        },
        {
          kind: 'item',
          id: 'show-version-history',
          label: 'Version history',
          icon: 'history',
          onClick: () => {
            if (!ui.historyPanelVisible) ui.toggleHistoryPanel();
          },
        },
        { kind: 'separator', id: 'sep-props' },
        {
          kind: 'item',
          id: 'properties',
          label: 'Properties…',
          icon: 'info',
          onClick: () => setShowProperties(true),
        },
        {
          kind: 'item',
          id: 'about',
          label: 'About casual sheets',
          icon: 'help_outline',
          onClick: () => setShowAbout(true),
        },
      ],
    },
    edit: {
      label: 'Edit',
      items: [
        { kind: 'item', id: 'undo', label: 'Undo', icon: 'undo', shortcut: 'Ctrl+Z', run: undo },
        { kind: 'item', id: 'redo', label: 'Redo', icon: 'redo', shortcut: 'Ctrl+Y', run: redo },
        { kind: 'separator', id: 'sep-clip' },
        {
          kind: 'item',
          id: 'cut',
          label: 'Cut',
          icon: 'content_cut',
          shortcut: 'Ctrl+X',
          run: actCut,
        },
        {
          kind: 'item',
          id: 'copy',
          label: 'Copy',
          icon: 'content_copy',
          shortcut: 'Ctrl+C',
          run: actCopy,
        },
        {
          kind: 'item',
          id: 'paste',
          label: 'Paste',
          icon: 'content_paste',
          shortcut: 'Ctrl+V',
          run: actPaste,
        },
        {
          kind: 'item',
          id: 'paste-format',
          label: 'Paste formatting only',
          icon: 'content_paste',
          shortcut: 'Ctrl+Shift+V',
          run: pasteFormattingOnly,
        },
        {
          kind: 'item',
          id: 'paste-special',
          label: 'Paste Special…',
          icon: 'content_paste_go',
          shortcut: 'Ctrl+Alt+V',
          onClick: () => setShowPasteSpecial(true),
        },
        { kind: 'separator', id: 'sep-find' },
        {
          kind: 'item',
          id: 'find-replace',
          label: 'Find & Replace…',
          icon: 'search',
          shortcut: 'Ctrl+F',
          run: openFindReplace,
        },
        {
          kind: 'item',
          id: 'go-to-special',
          label: 'Go To Special…',
          icon: 'highlight_alt',
          shortcut: 'Ctrl+G',
          onClick: () => setShowGoToSpecial(true),
        },
        { kind: 'separator', id: 'sep-cells' },
        // The Insert / Delete dialogs were keyboard-only via Polish #1;
        // surface them in the menu so they're discoverable.
        {
          kind: 'item',
          id: 'edit-insert-cells',
          label: 'Insert cells…',
          icon: 'add_box',
          shortcut: 'Ctrl++',
          onClick: () => setCellsOp('insert'),
        },
        {
          kind: 'item',
          id: 'edit-delete-cells',
          label: 'Delete cells…',
          icon: 'indeterminate_check_box',
          shortcut: 'Ctrl+-',
          onClick: () => setCellsOp('delete'),
        },
        { kind: 'separator', id: 'sep-sel' },
        {
          kind: 'item',
          id: 'edit-select-col',
          label: 'Select column',
          icon: 'view_column',
          shortcut: 'Ctrl+Space',
          onClick: () => api && selectEntireColumns(api),
        },
        {
          kind: 'item',
          id: 'edit-select-row',
          label: 'Select row',
          icon: 'view_stream',
          shortcut: 'Shift+Space',
          onClick: () => api && selectEntireRows(api),
        },
        {
          kind: 'item',
          id: 'edit-edit-cell',
          label: 'Edit cell',
          icon: 'edit',
          shortcut: 'F2',
          onClick: () => api && enterCellEditMode(api),
        },
      ],
    },
    view: {
      label: 'View',
      items: [
        {
          kind: 'item',
          id: 'toggle-formula-bar',
          label: ui.formulaBarVisible ? 'Hide formula bar' : 'Show formula bar',
          icon: 'functions',
          onClick: ui.toggleFormulaBar,
        },
        {
          kind: 'item',
          id: 'toggle-compact-ribbon',
          label: ui.ribbonCompact ? '✓ Compact ribbon' : 'Compact ribbon',
          icon: 'density_small',
          onClick: ui.toggleRibbonCompact,
        },
        {
          kind: 'item',
          id: 'show-formulas',
          label: ui.showFormulas ? '✓ Show formulas' : 'Show formulas',
          icon: 'description',
          shortcut: 'Ctrl+`',
          onClick: ui.toggleShowFormulas,
        },
        {
          kind: 'item',
          id: 'toggle-gridlines',
          label: 'Gridlines',
          icon: 'grid_on',
          onClick: () => api && toggleGridlines(api),
        },
        {
          kind: 'item',
          id: 'toggle-watermark',
          label: watermarkOn ? '✓ Watermark…' : 'Watermark…',
          icon: 'water_drop',
          onClick: () => {
            if (!api) return;
            setShowWatermark(true);
          },
        },
        { kind: 'separator', id: 'sep-freeze' },
        {
          kind: 'item',
          id: 'freeze-row',
          label: 'Freeze top row',
          icon: 'border_horizontal',
          run: freezeFirstRow,
        },
        {
          kind: 'item',
          id: 'freeze-col',
          label: 'Freeze first column',
          icon: 'border_vertical',
          run: freezeFirstColumn,
        },
        {
          kind: 'item',
          id: 'freeze-selection',
          label: 'Freeze panes (at selection)',
          icon: 'grid_4x4',
          run: freezeAtSelection,
        },
        { kind: 'item', id: 'unfreeze', label: 'Unfreeze', icon: 'grid_off', run: unfreezePanes },
        { kind: 'separator', id: 'sep-nav' },
        {
          kind: 'item',
          id: 'jump-home',
          label: 'Jump to A1',
          icon: 'home',
          shortcut: 'Ctrl+Home',
          onClick: () => api && jumpToFirstCell(api),
        },
        {
          kind: 'item',
          id: 'jump-end',
          label: 'Jump to last cell',
          icon: 'last_page',
          shortcut: 'Ctrl+End',
          onClick: () => api && jumpToLastCell(api),
        },
        {
          kind: 'item',
          id: 'prev-sheet',
          label: 'Previous sheet',
          icon: 'navigate_before',
          shortcut: 'Ctrl+PageUp',
          onClick: () => api && switchToPreviousSheet(api),
        },
        {
          kind: 'item',
          id: 'next-sheet',
          label: 'Next sheet',
          icon: 'navigate_next',
          shortcut: 'Ctrl+PageDown',
          onClick: () => api && switchToNextSheet(api),
        },
        { kind: 'separator', id: 'sep-panels' },
        {
          kind: 'item',
          id: 'tables-panel',
          label: ui.tablesPanelVisible ? 'Hide Tables panel' : 'Tables panel',
          icon: 'table_rows',
          onClick: ui.toggleTablesPanel,
        },
        {
          kind: 'item',
          id: 'outline-panel',
          label: ui.outlinePanelVisible ? 'Hide Outline panel' : 'Outline panel',
          icon: 'list',
          onClick: ui.toggleOutlinePanel,
        },
        {
          kind: 'item',
          id: 'charts-panel',
          label: ui.chartsPanelVisible ? 'Hide Charts panel' : 'Charts panel',
          icon: 'bar_chart',
          onClick: ui.toggleChartsPanel,
        },
        {
          kind: 'item',
          id: 'pivot-fields-panel',
          label: ui.pivotPanelVisible ? 'Hide PivotTable Fields' : 'PivotTable Fields',
          icon: 'pivot_table_chart',
          onClick: ui.togglePivotPanel,
        },
        {
          kind: 'item',
          id: 'history-panel',
          label: ui.historyPanelVisible ? 'Hide History panel' : 'History panel',
          icon: 'history',
          onClick: ui.toggleHistoryPanel,
        },
        {
          kind: 'item',
          id: 'comments-panel',
          label: 'Comments panel',
          icon: 'forum',
          run: toggleCommentPanel,
        },
      ],
    },
    insert: {
      label: 'Insert',
      items: [
        // High-leverage objects first — what an Excel user reaches for.
        {
          kind: 'item',
          id: 'new-sheet',
          label: 'New sheet',
          icon: 'add_box',
          shortcut: 'Shift+F11',
          run: insertNewSheet,
        },
        {
          kind: 'item',
          id: 'insert-table',
          label: 'Table',
          icon: 'table_rows',
          shortcut: 'Ctrl+T',
          run: insertTable,
        },
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
          id: 'insert-sparkline',
          label: 'Sparkline…',
          icon: 'show_chart',
          onClick: () => setShowInsertSparkline(true),
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
        {
          kind: 'item',
          id: 'insert-function',
          label: 'Function…',
          icon: 'functions',
          shortcut: 'Shift+F3',
          onClick: () => document.dispatchEvent(new CustomEvent('casual-open-insert-function')),
        },
        {
          kind: 'item',
          id: 'insert-link',
          label: 'Hyperlink…',
          icon: 'link',
          shortcut: 'Ctrl+K',
          run: insertHyperlink,
        },
        {
          kind: 'item',
          id: 'insert-comment',
          label: 'Comment',
          icon: 'comment',
          shortcut: 'Shift+F2',
          run: insertComment,
        },
        { kind: 'separator', id: 'sep-rowcol' },
        {
          kind: 'submenu',
          id: 'insert-rowcol',
          label: 'Rows & columns',
          icon: 'grid_on',
          items: [
            {
              kind: 'item',
              id: 'insert-row-above',
              label: 'Row above',
              icon: 'vertical_align_top',
              run: insertRowAbove,
            },
            {
              kind: 'item',
              id: 'insert-row-below',
              label: 'Row below',
              icon: 'vertical_align_bottom',
              run: insertRowBelow,
            },
            {
              kind: 'item',
              id: 'insert-col-left',
              label: 'Column left',
              icon: 'keyboard_tab_rtl',
              run: insertColumnLeft,
            },
            {
              kind: 'item',
              id: 'insert-col-right',
              label: 'Column right',
              icon: 'keyboard_tab',
              run: insertColumnRight,
            },
          ],
        },
        { kind: 'separator', id: 'sep-autofit' },
        {
          kind: 'item',
          id: 'autofit-col',
          label: 'Auto-fit column width',
          icon: 'settings_ethernet',
          run: autoFitColumns,
        },
        {
          kind: 'item',
          id: 'autofit-row',
          label: 'Auto-fit row height',
          icon: 'height',
          run: autoFitRows,
        },
        { kind: 'separator', id: 'sep-date' },
        {
          kind: 'item',
          id: 'insert-today',
          label: "Today's date",
          icon: 'today',
          shortcut: 'Ctrl+;',
          run: insertTodayDate,
        },
        {
          kind: 'item',
          id: 'insert-time',
          label: 'Current time',
          icon: 'schedule',
          shortcut: 'Ctrl+Shift+:',
          run: insertCurrentTime,
        },
      ],
    },
    format: {
      label: 'Format',
      items: [
        {
          kind: 'item',
          id: 'format-cells',
          label: 'Format cells…',
          icon: 'format_shapes',
          shortcut: 'Ctrl+1',
          onClick: () => setShowFormatCells(true),
        },
        {
          kind: 'submenu',
          id: 'cell-styles',
          label: 'Cell styles',
          icon: 'palette',
          items: [
            {
              kind: 'item',
              id: 'cell-style-normal',
              label: 'Normal',
              run: (a) => applyCellStyle(a, 'normal'),
            },
            {
              kind: 'item',
              id: 'cell-style-good',
              label: 'Good',
              run: (a) => applyCellStyle(a, 'good'),
            },
            {
              kind: 'item',
              id: 'cell-style-bad',
              label: 'Bad',
              run: (a) => applyCellStyle(a, 'bad'),
            },
            {
              kind: 'item',
              id: 'cell-style-neutral',
              label: 'Neutral',
              run: (a) => applyCellStyle(a, 'neutral'),
            },
            {
              kind: 'item',
              id: 'cell-style-title',
              label: 'Title',
              run: (a) => applyCellStyle(a, 'title'),
            },
            {
              kind: 'item',
              id: 'cell-style-heading1',
              label: 'Heading 1',
              run: (a) => applyCellStyle(a, 'heading1'),
            },
            {
              kind: 'item',
              id: 'cell-style-heading2',
              label: 'Heading 2',
              run: (a) => applyCellStyle(a, 'heading2'),
            },
          ],
        },
        { kind: 'separator', id: 'sep-format-cells' },
        // Number formats live behind a submenu so they don't push the
        // other Format actions off the bottom of the dropdown. The user
        // hovers "Number format" and gets all 10 variants in one place.
        // Excel's Ctrl+Shift+1..6 bindings map to a subset of these keys —
        // expose the shortcut hint on the items it lines up with.
        {
          kind: 'submenu',
          id: 'num-format',
          label: 'Number format',
          icon: 'looks_one',
          items: (
            [
              'general',
              'number',
              'integer',
              'currency',
              'accounting',
              'percent',
              'date',
              'time',
              'scientific',
              'text',
            ] as NumberFormatKey[]
          ).map<MenuItem>((k) => {
            const shortcut: Record<string, string> = {
              number: 'Ctrl+Shift+1',
              time: 'Ctrl+Shift+2',
              date: 'Ctrl+Shift+3',
              currency: 'Ctrl+Shift+4',
              percent: 'Ctrl+Shift+5',
              scientific: 'Ctrl+Shift+6',
            };
            return {
              kind: 'item',
              id: `num-${k}`,
              label: k[0]!.toUpperCase() + k.slice(1),
              icon: 'looks_one',
              shortcut: shortcut[k],
              onClick: () => api && setNumberFormatByKey(api, k),
            };
          }),
        },
        {
          kind: 'item',
          id: 'decimal-up',
          label: 'Increase decimals',
          icon: 'add',
          run: increaseDecimal,
        },
        {
          kind: 'item',
          id: 'decimal-down',
          label: 'Decrease decimals',
          icon: 'remove',
          run: decreaseDecimal,
        },
        { kind: 'separator', id: 'sep-cond' },
        {
          kind: 'item',
          id: 'conditional-formatting',
          label: 'Conditional formatting…',
          icon: 'palette',
          run: openConditionalFormatting,
        },
        { kind: 'separator', id: 'sep-visibility' },
        // Hide / Unhide grouped — Excel's Format → Visibility submenu.
        {
          kind: 'submenu',
          id: 'visibility',
          label: 'Visibility',
          icon: 'visibility',
          items: [
            {
              kind: 'item',
              id: 'hide-row',
              label: 'Hide row',
              icon: 'visibility_off',
              shortcut: 'Ctrl+9',
              run: hideSelectedRows,
            },
            {
              kind: 'item',
              id: 'unhide-row',
              label: 'Unhide row',
              icon: 'visibility',
              shortcut: 'Ctrl+Shift+9',
              run: unhideSelectedRows,
            },
            {
              kind: 'item',
              id: 'hide-col',
              label: 'Hide column',
              icon: 'visibility_off',
              shortcut: 'Ctrl+0',
              run: hideSelectedColumns,
            },
            {
              kind: 'item',
              id: 'unhide-col',
              label: 'Unhide column',
              icon: 'visibility',
              shortcut: 'Ctrl+Shift+0',
              run: unhideSelectedColumns,
            },
          ],
        },
        { kind: 'separator', id: 'sep-fit' },
        {
          kind: 'item',
          id: 'autofit-col',
          label: 'Auto-fit column width',
          icon: 'settings_ethernet',
          run: autoFitColumns,
        },
        {
          kind: 'item',
          id: 'autofit-row',
          label: 'Auto-fit row height',
          icon: 'height',
          run: autoFitRows,
        },
        { kind: 'separator', id: 'sep-delete' },
        {
          kind: 'item',
          id: 'delete-row',
          label: 'Delete row',
          icon: 'delete_sweep',
          run: deleteSelectedRow,
        },
        {
          kind: 'item',
          id: 'delete-col',
          label: 'Delete column',
          icon: 'folder_delete',
          run: deleteSelectedColumn,
        },
      ],
    },
    data: {
      label: 'Data',
      items: [
        {
          kind: 'item',
          id: 'sort-custom',
          label: 'Sort range…',
          icon: 'sort',
          run: openCustomSort,
        },
        {
          kind: 'item',
          id: 'sort-custom-list',
          label: 'Sort by custom list…',
          icon: 'format_list_numbered',
          onClick: () => setShowCustomListSort(true),
        },
        {
          kind: 'item',
          id: 'sort-by-color',
          label: 'Sort by colour…',
          icon: 'palette',
          onClick: () => setShowColorSort(true),
        },
        {
          kind: 'item',
          id: 'data-validation',
          label: 'Data validation…',
          icon: 'rule',
          run: openDataValidation,
        },
        {
          kind: 'item',
          id: 'protect-sheet',
          label: protectedOn ? '✓ Make workbook read-only' : 'Make workbook read-only',
          icon: 'lock',
          onClick: toggleProtect,
        },
        {
          kind: 'item',
          id: 'protect-worksheet',
          label: sheetProtectedOn ? '✓ Protect sheet' : 'Protect sheet',
          icon: 'shield',
          onClick: toggleSheetProtect,
        },
        {
          kind: 'item',
          id: 'protect-range',
          label: 'Protect range',
          icon: 'lock_person',
          onClick: handleProtectRange,
        },
        {
          kind: 'item',
          id: 'remove-range-protection',
          label: 'Remove range protection',
          icon: 'lock_open',
          onClick: handleRemoveRangeProtections,
        },
        {
          kind: 'item',
          id: 'name-manager',
          label: 'Name Manager…',
          icon: 'bookmark_add',
          shortcut: 'Ctrl+F3',
          onClick: () => setShowNameManager(true),
        },
        {
          kind: 'item',
          id: 'goal-seek',
          label: 'Goal Seek…',
          icon: 'analytics',
          onClick: () => setShowGoalSeek(true),
        },
        {
          kind: 'item',
          id: 'scenario-manager',
          label: 'Scenario Manager…',
          icon: 'tune',
          onClick: () => setShowScenarioManager(true),
        },
        {
          kind: 'item',
          id: 'iterative-calc',
          label: iterativeCalc ? '✓ Iterative calculation' : 'Iterative calculation',
          icon: 'sync',
          onClick: () => {
            if (!api) return;
            const next = !iterativeCalc;
            setIterativeCalc(next);
            // Excel's "Enable iterative calculation": >1 cycle count lets
            // circular references converge instead of erroring. 100 matches
            // Excel's default max iterations; 1 disables iteration.
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            (api as any).getFormula?.()?.setMaxIteration?.(next ? 100 : 1);
            forceRecalculate(api);
          },
        },
        {
          kind: 'submenu',
          id: 'macros',
          label: 'Macros',
          icon: 'smart_button',
          items: [
            {
              kind: 'item',
              id: 'macro-record',
              label: macroRecording ? 'Stop recording' : 'Record macro',
              icon: macroRecording ? 'stop_circle' : 'fiber_manual_record',
              onClick: macroRecording ? handleStopMacro : handleRecordMacro,
            },
            ...(listMacros().length > 0
              ? ([
                  { kind: 'separator', id: 'sep-macros' },
                  ...listMacros().map((m) => ({
                    kind: 'item' as const,
                    id: `macro-run-${m.name.replace(/\s+/g, '-')}`,
                    label: `Run "${m.name}"`,
                    icon: 'play_arrow',
                    shortcut: m.shortcut ? `Ctrl+Shift+${m.shortcut}` : undefined,
                    onClick: () => handleRunMacro(m.name),
                  })),
                  { kind: 'separator', id: 'sep-macros-manage' },
                  {
                    kind: 'item',
                    id: 'macro-manage',
                    label: 'Manage macros…',
                    icon: 'settings',
                    onClick: () => setShowMacros(true),
                  },
                ] as MenuItem[])
              : []),
          ],
        },
        {
          kind: 'item',
          id: 'flash-fill',
          label: 'Flash Fill',
          icon: 'auto_awesome',
          shortcut: 'Ctrl+E',
          run: runFlashFillWithToast,
        },
        {
          kind: 'item',
          id: 'refresh-pivots',
          label: 'Refresh PivotTables',
          icon: 'autorenew',
          onClick: () => {
            if (!api) return;
            for (const p of pivots.pivots) {
              const extent = refreshPivot(api, p);
              if (extent) pivots.update(p.id, { lastOutputExtent: extent });
            }
          },
        },
        {
          kind: 'item',
          id: 'drill-down',
          label: 'Drill down (selected pivot cell)',
          icon: 'open_in_new',
          shortcut: 'Ctrl+Shift+D',
          onClick: () => api && runDrillDown(),
        },
        { kind: 'separator', id: 'sep-clean' },
        {
          kind: 'item',
          id: 'text-to-columns',
          label: 'Text to Columns…',
          icon: 'splitscreen',
          onClick: () => setShowTextToColumns(true),
        },
        {
          kind: 'item',
          id: 'remove-duplicates',
          label: 'Remove Duplicates…',
          icon: 'filter_list_off',
          onClick: () => setShowRemoveDuplicates(true),
        },
        {
          kind: 'item',
          id: 'subtotal',
          label: 'Subtotal…',
          icon: 'functions',
          onClick: () => setShowSubtotals(true),
        },
        {
          kind: 'item',
          id: 'advanced-filter',
          label: 'Advanced Filter…',
          icon: 'filter_alt',
          onClick: () => setShowAdvancedFilter(true),
        },
        {
          kind: 'item',
          id: 'show-all-rows',
          label: 'Show all rows',
          icon: 'unfold_more',
          run: showAllRows,
        },
        { kind: 'separator', id: 'sep-outline' },
        {
          kind: 'item',
          id: 'group-rows',
          label: 'Group rows',
          icon: 'unfold_less',
          onClick: () => {
            outlineActions.groupRows();
          },
        },
        {
          kind: 'item',
          id: 'group-cols',
          label: 'Group columns',
          icon: 'view_week',
          onClick: () => {
            outlineActions.groupCols();
          },
        },
        {
          kind: 'item',
          id: 'ungroup',
          label: 'Ungroup',
          icon: 'unfold_more_double',
          onClick: () => {
            outlineActions.ungroupSelection();
          },
        },
        { kind: 'separator', id: 'sep-formula-audit' },
        {
          kind: 'item',
          id: 'evaluate-formula',
          label: 'Evaluate Formula…',
          icon: 'troubleshoot',
          onClick: () => setShowEvaluateFormula(true),
        },
        {
          kind: 'item',
          id: 'watch-window',
          label: ui.watchPanelVisible ? 'Hide Watch Window' : 'Watch Window',
          icon: 'visibility',
          onClick: ui.toggleWatchPanel,
        },
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
        {
          kind: 'item',
          id: 'command-search',
          label: 'Search / Tell Me…',
          icon: 'search',
          // Alt+Q is the Office "Tell Me" chord; we also bind
          // Ctrl+Shift+P (VS Code / Linear / Notion command palette).
          // Showing Alt+Q here since it's the more discoverable one
          // for users coming from Excel — the cheat sheet lists both.
          shortcut: 'Alt+Q',
          onClick: () => setShowCommandSearch(true),
        },
        {
          kind: 'item',
          id: 'keyboard-shortcuts',
          label: 'Keyboard shortcuts',
          icon: 'info',
          shortcut: 'Ctrl+/',
          onClick: () => setShowShortcuts(true),
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
          onClick: () => openExternal('https://github.com/CasualOffice/sheets'),
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
                // `run` may be async (lazy-plugin loaders await
                // ensurePluginByName before dispatching). Fire-and-forget the
                // promise — but DO invoke it so its async work runs; previously
                // the result was discarded synchronously, which still ran sync
                // actions but is brittle. Keep onClose immediate so the menu
                // closes responsively while the action resolves.
                if (item.run && api) void item.run(api);
                if (item.onClick) item.onClick();
                onClose();
              }}
            />
          </MenuItemButton>
        ))}
      </div>

      {showProperties && <PropertiesDialog onClose={() => setShowProperties(false)} />}

      {showFormatCells && <FormatCellsDialog onClose={() => setShowFormatCells(false)} />}

      {showAbout && <AboutDialog onClose={() => setShowAbout(false)} />}

      {showShortcuts && <KeyboardShortcutsDialog onClose={() => setShowShortcuts(false)} />}

      {showCommandSearch && (
        <CommandSearchDialog
          items={collectCommandSearchItems(
            Object.values(menus).flatMap((menu) => menu.items),
            api,
          )}
          onClose={() => setShowCommandSearch(false)}
        />
      )}

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
              // Transient confirmation — the chart appears as a
              // floating overlay over the grid, which a user
              // looking at the source range can miss entirely. The
              // toast confirms the action AND announces the new
              // chart's name so the user can find it in the
              // Charts panel.
              toast.success(`Added ${name}`);
            } else {
              // buildChartModelForRange returns null when the
              // range can't be coerced (e.g. all-empty selection).
              // The dialog's pre-flight catches most of these but
              // some edge cases fall through silently — surface
              // them via toast so the user isn't left wondering
              // why nothing appeared.
              toast.error("Couldn't build a chart from that range — check the source data");
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
          onConfirm={({
            source,
            target,
            rowFieldColumns,
            rowGrouping,
            colFieldColumns,
            valueFields,
            filters,
          }) => {
            const wb = api.getActiveWorkbook();
            const ws = activeSheet(api);
            if (!wb || !ws) {
              setShowInsertPivot(false);
              return;
            }
            const sheetId = facadeSheetId(ws);
            if (!sheetId) {
              setShowInsertPivot(false);
              return;
            }
            const model = {
              id: newPivotId(),
              sourceSheetId: sheetId,
              source,
              targetSheetId: sheetId,
              target,
              // Column indices in the dialog are relative to the source range's
              // left edge; the model stores absolute column offsets within the
              // range too. The dialog already gives us the in-range index, which
              // matches what compute.ts expects. The array is outer-first; an
              // empty array means Grand-Total-only.
              // Date grouping applies to the primary (outermost) row field.
              rows: rowFieldColumns.map((column, i) => ({
                column,
                grouping: i === 0 ? rowGrouping : undefined,
              })),
              // A non-empty cols list switches compute.ts to the cross-tab
              // / matrix layout (value fans out one block per distinct
              // column-field value). Empty keeps the classic row layout.
              cols: colFieldColumns.map((column) => ({ column })),
              values: valueFields.map((v) => ({
                column: v.column,
                agg: v.aggregation,
                showAs: v.showAs,
              })),
              filters,
              title: `PivotTable ${pivots.pivots.length + 1}`,
            };
            const extent = applyPivot(api, model);
            // Persist the extent on the model so a later refresh can
            // clear the previous output before writing the new one.
            pivots.insert(extent ? { ...model, lastOutputExtent: extent } : model);
            setShowInsertPivot(false);
            // Excel opens the PivotTable Fields list right after insert so
            // you can refine the layout without hunting for the pane.
            if (!ui.pivotPanelVisible) ui.togglePivotPanel();
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
            if (!api) return;
            const result = printActiveSheet(api, options);
            if (result.ok) return;
            // Soft-guard refusal (#50): one big HTML table OOMs the
            // tab past ~100k cells, so we surface a clear next step
            // instead of crashing. The "Set Print Area" pointer is
            // recoverable in a couple of clicks (the menu item already
            // exists right below Print).
            if (result.reason === 'too-large') {
              toast.error(
                `Too many cells to print at once (${result.cellCount.toLocaleString()} > ` +
                  `${result.limit.toLocaleString()} limit). Try File → Set Print Area to ` +
                  `pick a smaller range, or split the sheet.`,
                { duration: 12_000 },
              );
            } else {
              toast.error('Nothing to print — the active sheet looks empty.');
            }
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

      {showPasteSpecial && (
        <PasteSpecialDialog
          onCancel={() => setShowPasteSpecial(false)}
          onConfirm={(mode: PasteSpecialMode) => {
            setShowPasteSpecial(false);
            if (!api) return;
            pasteSpecial(api, mode);
            // "Formats" / "Column widths" / "All except borders"
            // produce changes that are easy to miss at a glance
            // (no new cell text, only style or geometry tweaks).
            // A short confirmation pins the action so the user
            // knows the variant they picked actually ran.
            toast.success(`Pasted: ${PASTE_SPECIAL_LABEL[mode]}`);
          }}
        />
      )}

      {showNameManager && api && (
        <NameManagerDialog api={api} onClose={() => setShowNameManager(false)} />
      )}

      {showGoalSeek && api && <GoalSeekDialog api={api} onClose={() => setShowGoalSeek(false)} />}
      {showScenarioManager && api && (
        <ScenarioManagerDialog api={api} onClose={() => setShowScenarioManager(false)} />
      )}
      {showGoToSpecial && api && (
        <GoToSpecialDialog api={api} onClose={() => setShowGoToSpecial(false)} />
      )}
      {showRemoveDuplicates && api && (
        <RemoveDuplicatesDialog api={api} onClose={() => setShowRemoveDuplicates(false)} />
      )}
      {showCustomListSort && api && (
        <CustomListSortDialog api={api} onClose={() => setShowCustomListSort(false)} />
      )}
      {showColorSort && api && (
        <ColorSortDialog api={api} onClose={() => setShowColorSort(false)} />
      )}
      {showEvaluateFormula && api && (
        <EvaluateFormulaDialog api={api} onClose={() => setShowEvaluateFormula(false)} />
      )}
      {showTextToColumns && api && (
        <TextToColumnsDialog api={api} onClose={() => setShowTextToColumns(false)} />
      )}
      {showSubtotals && api && (
        <SubtotalsDialog api={api} onClose={() => setShowSubtotals(false)} />
      )}
      {showAdvancedFilter && api && (
        <AdvancedFilterDialog api={api} onClose={() => setShowAdvancedFilter(false)} />
      )}
      {showMacros && api && (
        <MacrosDialog
          api={api}
          onClose={() => {
            setShowMacros(false);
            setMacroTick((t) => t + 1);
          }}
          onRan={(name, n) => toast.success(`Ran ${name} (${n} step${n === 1 ? '' : 's'})`)}
        />
      )}

      {showWatermark && api && (
        <WatermarkDialog
          api={api}
          initialOn={watermarkOn}
          initial={watermarkConfig}
          onClose={() => setShowWatermark(false)}
          onApplied={(on) => {
            setWatermarkOn(on);
            // Re-read the just-persisted preference so the dialog re-opens
            // with the user's latest text/opacity even after a clear.
            setWatermarkConfig(loadWatermarkPref());
          }}
        />
      )}

      {drillDownResult && (
        <DrillDownDialog result={drillDownResult} onClose={() => setDrillDownResult(null)} />
      )}

      {showInsertSparkline && api && (
        <InsertSparklineDialog
          api={api}
          defaultSourceA1={(() => {
            const sel = getActiveSelectionRange(api);
            return sel ? rangeToA1(sel) : '';
          })()}
          onCancel={() => setShowInsertSparkline(false)}
          onConfirm={({ type, source, anchor }) => {
            setShowInsertSparkline(false);
            const wb = api.getActiveWorkbook();
            const ws = activeSheet(api);
            const sheetId = ws ? facadeSheetId(ws) : null;
            if (!wb || !ws || !sheetId) {
              toast.error("Couldn't add sparkline: no active sheet");
              return;
            }
            try {
              sparklinesCtx.add({
                type,
                unitId: wb.getId(),
                sheetId,
                source,
                anchor,
              });
              // Capitalise type for the readable label ("line" → "Line").
              toast.success(
                `Added ${type.charAt(0).toUpperCase() + type.slice(1)} sparkline at ${anchor}`,
              );
            } catch (err) {
              const msg = err instanceof Error ? err.message : String(err);
              toast.error(`Couldn't add sparkline: ${msg}`);
            }
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
  // Cap height + scroll when even a top-anchored menu is taller than the
  // viewport (otherwise its lower items render off-screen, unclickable).
  const [maxH, setMaxH] = useState<number | null>(null);

  useEffect(() => {
    if (!isOpen) return;
    const rect = btnRef.current?.getBoundingClientRect();
    if (rect) setPos({ top: rect.bottom + 2, left: rect.left });
  }, [isOpen]);

  // Bottom-anchor a tall menu so it stays inside the viewport (the menu has no
  // scroll — that would clip the fly-out submenus — so instead we shift it up
  // when it would overflow the bottom edge). Without this, the lower items of a
  // long menu (e.g. Data) render below the viewport and can't be clicked.
  useLayoutEffect(() => {
    if (!isOpen) return;
    const pop = popRef.current;
    if (!pop) return;
    const margin = 8;
    const avail = window.innerHeight - margin * 2;
    // scrollHeight is the full content height even once maxHeight clips it.
    const h = pop.scrollHeight;
    if (h > avail) {
      // Taller than the viewport: pin to the top and let it scroll so every
      // item is reachable (submenus still fly out to the side).
      setMaxH(avail);
      setPos((cur) => (cur.top === margin ? cur : { ...cur, top: margin }));
      return;
    }
    setMaxH(null);
    setPos((cur) => {
      if (cur.top + h <= window.innerHeight - margin) return cur;
      const top = Math.max(margin, window.innerHeight - h - margin);
      return top === cur.top ? cur : { ...cur, top };
    });
  }, [isOpen, pos.top]);

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
          style={{
            top: pos.top,
            left: pos.left,
            ...(maxH != null ? { maxHeight: maxH, overflowY: 'auto' as const } : null),
          }}
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
            {item.shortcut && (
              <span className="menu__item-shortcut">
                {formatShortcut(item.shortcut, navigator.platform)}
              </span>
            )}
          </button>
        );
      })}
    </>
  );
}
