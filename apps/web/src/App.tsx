import { useCallback, useMemo, useState, type ReactNode } from 'react';
import type { IWorkbookData } from '@univerjs/core';
import { TitleBar } from './shell/TitleBar';
import { MenuBar } from './shell/MenuBar';
import { Toolbar } from './shell/Toolbar';
import { FormulaBar } from './shell/FormulaBar';
import { SheetTabs } from './shell/SheetTabs';
import { TablesPanel } from './shell/TablesPanel';
import { UniverSheet } from './UniverSheet';
import { emptyWorkbook } from './snapshot';
import { UniverRoot } from './UniverRoot';
import { useWorkbookGrowth } from './hooks/useWorkbookGrowth';
import { useFileDrop } from './hooks/useFileDrop';
import { WorkbookContext, type WorkbookCtxValue, type WorkbookFormat } from './workbook-context';
import { UIContext, type UICtxValue } from './ui-context';

export function App() {
  const [snapshot, setSnapshot] = useState<IWorkbookData>(() => emptyWorkbook());
  const [formulaBarVisible, setFormulaBarVisible] = useState(true);
  const [tablesPanelVisible, setTablesPanelVisible] = useState(false);
  // What File → Save should write to. Set on Open from the file extension;
  // null while editing an empty / unsaved workbook (in which case Save
  // defaults to .xlsx).
  const [sourceFormat, setSourceFormat] = useState<WorkbookFormat | null>(null);

  const replaceWorkbook = useCallback(
    (next: IWorkbookData, format?: WorkbookFormat | null) => {
      setSnapshot(next);
      if (format !== undefined) setSourceFormat(format);
    },
    [],
  );

  const wbValue: WorkbookCtxValue = useMemo(
    () => ({ snapshot, replaceWorkbook, sourceFormat }),
    [snapshot, replaceWorkbook, sourceFormat],
  );

  const uiValue: UICtxValue = useMemo(
    () => ({
      formulaBarVisible,
      toggleFormulaBar: () => setFormulaBarVisible((v) => !v),
      tablesPanelVisible,
      toggleTablesPanel: () => setTablesPanelVisible((v) => !v),
    }),
    [formulaBarVisible, tablesPanelVisible],
  );

  return (
    <UniverRoot>
      <UIContext.Provider value={uiValue}>
        <WorkbookContext.Provider value={wbValue}>
          <GrowthDriver />
          <FileDropDriver />
          <div
            className={`app${formulaBarVisible ? '' : ' app--no-formula-bar'}`}
            data-testid="app-shell"
          >
            <TitleBar />
            <MenuBar />
            <Toolbar />
            {formulaBarVisible && <FormulaBar />}
            <div className="grid-row">
              <main className="grid-host" data-testid="grid-host">
                <UniverSheet snapshot={snapshot} />
              </main>
              {tablesPanelVisible && <TablesPanel />}
            </div>
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

/** Window-level file drag-and-drop. Renders an overlay while a file is over
 * the page; the hook itself handles the actual drop and routes through the
 * shared open flow. */
function FileDropDriver(): ReactNode {
  const dragging = useFileDrop();
  if (!dragging) return null;
  return (
    <div className="file-drop-overlay" data-testid="file-drop-overlay" aria-hidden="true">
      <div className="file-drop-overlay__card">
        <div className="file-drop-overlay__title">Drop to open</div>
        <div className="file-drop-overlay__hint">.xlsx · .ods · .csv · .tsv</div>
      </div>
    </div>
  );
}
