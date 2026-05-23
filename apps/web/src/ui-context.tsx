import { createContext } from 'react';

export type UICtxValue = {
  formulaBarVisible: boolean;
  toggleFormulaBar: () => void;
  tablesPanelVisible: boolean;
  toggleTablesPanel: () => void;
  outlinePanelVisible: boolean;
  toggleOutlinePanel: () => void;
  chartsPanelVisible: boolean;
  toggleChartsPanel: () => void;
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
