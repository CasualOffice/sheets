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
  /** Show the "Share for co-editing" dialog. Lifted to app scope so the
   *  titlebar's primary Share button can open it without coupling to MenuBar. */
  openShareRoom: () => void;
};

export const UIContext = createContext<UICtxValue | null>(null);
