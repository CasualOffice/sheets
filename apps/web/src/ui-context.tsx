import { createContext } from 'react';

export type UICtxValue = {
  formulaBarVisible: boolean;
  toggleFormulaBar: () => void;
  /** Compact ribbon view (View → Compact ribbon) — Google-Sheets-style
   *  single-row toolbar. Opt-in, persisted; Full (2-row) is the default. */
  ribbonCompact: boolean;
  toggleRibbonCompact: () => void;
  tablesPanelVisible: boolean;
  toggleTablesPanel: () => void;
  outlinePanelVisible: boolean;
  toggleOutlinePanel: () => void;
  chartsPanelVisible: boolean;
  toggleChartsPanel: () => void;
  /** PivotTable Fields task pane — reconfigure a pivot's Filters / Columns
   *  / Rows / Values zones and re-apply it live (Excel's field pane). */
  pivotPanelVisible: boolean;
  togglePivotPanel: () => void;
  /** Comments task pane — our React panel that indexes thread comments on
   *  the active sheet (replaces Univer's bespoke comment sidebar so it
   *  shares the shared `.side-panel` shell + motion). */
  commentsPanelVisible: boolean;
  toggleCommentsPanel: () => void;
  /** Live session-history panel — read-only list of every mutation in
   *  the active room's Yjs op-log, who issued it, and when. */
  historyPanelVisible: boolean;
  toggleHistoryPanel: () => void;
  /** Imperative reset — closes every React side panel. Used by the
   *  PanelMutex when Univer's own sidebar opens (Comments) so two
   *  panels don't fight for the right edge. */
  closeAllReactPanels: () => void;
  /** Excel-style "Show Formulas" mode (Ctrl+`). When on, the
   *  ShowFormulasLayer paints formula source text over every cell
   *  that carries a formula. Toggle is non-destructive — turning it
   *  off restores normal rendering. */
  showFormulas: boolean;
  toggleShowFormulas: () => void;
  /** Show the "Share for co-editing" dialog. Lifted to app scope so the
   *  titlebar's primary Share button can open it without coupling to MenuBar. */
  openShareRoom: () => void;
};

export const UIContext = createContext<UICtxValue | null>(null);
