import { createContext } from 'react';

export type UICtxValue = {
  formulaBarVisible: boolean;
  toggleFormulaBar: () => void;
  tablesPanelVisible: boolean;
  toggleTablesPanel: () => void;
  outlinePanelVisible: boolean;
  toggleOutlinePanel: () => void;
};

export const UIContext = createContext<UICtxValue | null>(null);
