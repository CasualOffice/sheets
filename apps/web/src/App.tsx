import { useCallback, useState, type ReactNode } from 'react';
import type { IWorkbookData } from '@univerjs/core';
import { TitleBar } from './shell/TitleBar';
import { Ribbon } from './shell/Ribbon';
import { FormulaBar } from './shell/FormulaBar';
import { StatusBar } from './shell/StatusBar';
import { UniverSheet } from './UniverSheet';
import { emptyWorkbook } from './snapshot';
import { UniverRoot } from './UniverRoot';
import { useWorkbookGrowth } from './hooks/useWorkbookGrowth';
import { WorkbookContext, type WorkbookCtxValue } from './workbook-context';

export function App() {
  const [snapshot, setSnapshot] = useState<IWorkbookData>(() => emptyWorkbook());

  const replaceWorkbook = useCallback((next: IWorkbookData) => {
    setSnapshot(next);
  }, []);

  const ctxValue: WorkbookCtxValue = { snapshot, replaceWorkbook };

  return (
    <UniverRoot>
      <WorkbookContext.Provider value={ctxValue}>
        <GrowthDriver />
        <div className="app">
          <TitleBar filename={snapshot.name || 'Untitled'} />
          <Ribbon />
          <FormulaBar />
          <main className="grid-host" data-testid="grid-host">
            <UniverSheet snapshot={snapshot} />
          </main>
          <StatusBar />
        </div>
      </WorkbookContext.Provider>
    </UniverRoot>
  );
}

/** Effect-only component — auto-grows the active sheet near edges. */
function GrowthDriver(): ReactNode {
  useWorkbookGrowth();
  return null;
}
