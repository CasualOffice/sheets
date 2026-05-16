import { useCallback, useMemo, useState, type ReactNode } from 'react';
import type { IWorkbookData } from '@univerjs/core';
import { TitleBar } from './shell/TitleBar';
import { MenuBar } from './shell/MenuBar';
import { Toolbar } from './shell/Toolbar';
import { FormulaBar } from './shell/FormulaBar';
import { SheetTabs } from './shell/SheetTabs';
import { UniverSheet } from './UniverSheet';
import { emptyWorkbook } from './snapshot';
import { UniverRoot } from './UniverRoot';
import { useWorkbookGrowth } from './hooks/useWorkbookGrowth';
import { WorkbookContext, type WorkbookCtxValue } from './workbook-context';
import { UIContext, type UICtxValue } from './ui-context';

export function App() {
  const [snapshot, setSnapshot] = useState<IWorkbookData>(() => emptyWorkbook());
  const [formulaBarVisible, setFormulaBarVisible] = useState(true);

  const replaceWorkbook = useCallback((next: IWorkbookData) => {
    setSnapshot(next);
  }, []);

  const wbValue: WorkbookCtxValue = useMemo(
    () => ({ snapshot, replaceWorkbook }),
    [snapshot, replaceWorkbook],
  );

  const uiValue: UICtxValue = useMemo(
    () => ({
      formulaBarVisible,
      toggleFormulaBar: () => setFormulaBarVisible((v) => !v),
    }),
    [formulaBarVisible],
  );

  return (
    <UniverRoot>
      <UIContext.Provider value={uiValue}>
        <WorkbookContext.Provider value={wbValue}>
          <GrowthDriver />
          <div
            className={`app${formulaBarVisible ? '' : ' app--no-formula-bar'}`}
            data-testid="app-shell"
          >
            <TitleBar />
            <MenuBar />
            <Toolbar />
            {formulaBarVisible && <FormulaBar />}
            <main className="grid-host" data-testid="grid-host">
              <UniverSheet snapshot={snapshot} />
            </main>
            <SheetTabs />
          </div>
        </WorkbookContext.Provider>
      </UIContext.Provider>
    </UniverRoot>
  );
}

/** Effect-only component — auto-grows the active sheet near edges. */
function GrowthDriver(): ReactNode {
  useWorkbookGrowth();
  return null;
}
