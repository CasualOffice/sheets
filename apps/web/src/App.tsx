import { useMemo } from 'react';
import { TitleBar } from './shell/TitleBar';
import { Ribbon } from './shell/Ribbon';
import { StatusBar } from './shell/StatusBar';
import { UniverSheet } from './UniverSheet';
import { emptyWorkbook } from './snapshot';
import { UniverRoot } from './UniverRoot';

export function App() {
  const snapshot = useMemo(() => emptyWorkbook(), []);

  return (
    <UniverRoot>
      <div className="app">
        <TitleBar filename="Untitled" />
        <Ribbon />
        <main className="grid-host" data-testid="grid-host">
          <UniverSheet snapshot={snapshot} />
        </main>
        <StatusBar />
      </div>
    </UniverRoot>
  );
}
