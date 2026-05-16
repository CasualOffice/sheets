import { useMemo } from 'react';
import { TitleBar } from './shell/TitleBar';
import { Ribbon } from './shell/Ribbon';
import { FormulaBar } from './shell/FormulaBar';
import { StatusBar } from './shell/StatusBar';
import { UniverSheet } from './UniverSheet';
import { emptyWorkbook } from './snapshot';
import { UniverRoot } from './UniverRoot';
import { useWorkbookGrowth } from './hooks/useWorkbookGrowth';

export function App() {
  const snapshot = useMemo(() => emptyWorkbook(), []);

  return (
    <UniverRoot>
      <GrowthDriver />
      <div className="app">
        <TitleBar filename="Untitled" />
        <Ribbon />
        <FormulaBar />
        <main className="grid-host" data-testid="grid-host">
          <UniverSheet snapshot={snapshot} />
        </main>
        <StatusBar />
      </div>
    </UniverRoot>
  );
}

/** Effect-only component — auto-grows the active sheet near edges. */
function GrowthDriver() {
  useWorkbookGrowth();
  return null;
}
