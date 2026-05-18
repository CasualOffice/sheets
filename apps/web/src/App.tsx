import { useCallback, useMemo, useRef, useState, type ReactNode } from 'react';
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
import {
  WorkbookContext,
  type WorkbookCtxValue,
  type WorkbookFormat,
  type WorkbookMeta,
} from './workbook-context';
import { UIContext, type UICtxValue } from './ui-context';
import { OutlineProvider } from './outline/outline-context';
import { OutlinePanel } from './shell/OutlinePanel';
import { CollabDriver } from './collab/CollabDriver';
import { CreateRoomDialog } from './shell/CreateRoomDialog';
import { LoadingOverlay } from './shell/LoadingOverlay';
import { LoadingContext, type LoadingCtxValue, type LoadingState } from './loading-context';

export function App() {
  // Snapshot lives in a ref, NOT React state — see workbook-context.tsx.
  // Stage 3 of the large-file pipeline: keeping a multi-MB IWorkbookData
  // tree in React state alongside Univer's own copy doubled the peak heap
  // on big files. The ref carries the snapshot for the brief window
  // between replaceWorkbook and UniverSheet's swap effect, then it's
  // cleared so the data becomes GC-eligible.
  const initial = useMemo(() => emptyWorkbook(), []);
  const snapshotRef = useRef<IWorkbookData | null>(initial);
  const [meta, setMeta] = useState<WorkbookMeta>(() => ({
    id: initial.id ?? `wb-${Date.now()}`,
    name: initial.name ?? 'Untitled',
    sourceFormat: null,
    revision: 0,
  }));

  const [formulaBarVisible, setFormulaBarVisible] = useState(true);
  const [tablesPanelVisible, setTablesPanelVisible] = useState(false);
  const [outlinePanelVisible, setOutlinePanelVisible] = useState(false);
  const [shareRoomOpen, setShareRoomOpen] = useState(false);
  const [loading, setLoading] = useState<LoadingState | null>(null);

  const replaceWorkbook = useCallback(
    (next: IWorkbookData, format?: WorkbookFormat | null) => {
      snapshotRef.current = next;
      setMeta((prev) => ({
        id: next.id ?? prev.id,
        name: next.name ?? prev.name,
        sourceFormat: format !== undefined ? format : prev.sourceFormat,
        revision: prev.revision + 1,
      }));
      // Free the ref after consumers have processed the revision bump.
      // We wait two macrotasks: the first lets React flush its render +
      // useEffect pass (UniverSheet's swap, OutlineProvider's rehydrate);
      // the second is paranoia for any deferred work scheduled inside
      // those effects.
      setTimeout(() => {
        setTimeout(() => {
          if (snapshotRef.current === next) snapshotRef.current = null;
        }, 0);
      }, 0);
    },
    [],
  );

  // App owns only the meta update — TitleBar mirrors into Univer via
  // setName since App itself is outside the UniverProvider.
  const renameWorkbook = useCallback((name: string) => {
    const trimmed = name.trim();
    if (!trimmed) return;
    setMeta((prev) => (prev.name === trimmed ? prev : { ...prev, name: trimmed }));
  }, []);

  const wbValue: WorkbookCtxValue = useMemo(
    () => ({ meta, snapshotRef, replaceWorkbook, renameWorkbook }),
    [meta, replaceWorkbook, renameWorkbook],
  );

  const loadingValue: LoadingCtxValue = useMemo(
    () => ({
      state: loading,
      set: (next) => {
        if (next === null) {
          setLoading(null);
          return;
        }
        setLoading((prev) => {
          if (!prev) {
            if (!next.fileName || !next.phase) return prev;
            return {
              fileName: next.fileName,
              phase: next.phase,
              sizeBytes: next.sizeBytes,
              startedAt: Date.now(),
              error: next.error,
            };
          }
          return { ...prev, ...next, startedAt: prev.startedAt };
        });
      },
    }),
    [loading],
  );

  const uiValue: UICtxValue = useMemo(
    () => ({
      formulaBarVisible,
      toggleFormulaBar: () => setFormulaBarVisible((v) => !v),
      tablesPanelVisible,
      toggleTablesPanel: () => setTablesPanelVisible((v) => !v),
      outlinePanelVisible,
      toggleOutlinePanel: () => setOutlinePanelVisible((v) => !v),
      openShareRoom: () => setShareRoomOpen(true),
    }),
    [formulaBarVisible, tablesPanelVisible, outlinePanelVisible],
  );

  return (
    <UniverRoot>
      <UIContext.Provider value={uiValue}>
        <WorkbookContext.Provider value={wbValue}>
        <LoadingContext.Provider value={loadingValue}>
          <OutlineProvider>
            <GrowthDriver />
            <FileDropDriver />
            <CollabDriver>
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
                    <UniverSheet revision={meta.revision} initialSnapshot={initial} />
                  </main>
                  {tablesPanelVisible && <TablesPanel />}
                  {outlinePanelVisible && <OutlinePanel />}
                </div>
                <SheetTabs />
                {shareRoomOpen && (
                  <CreateRoomDialog onClose={() => setShareRoomOpen(false)} />
                )}
              </div>
            </CollabDriver>
            <LoadingOverlay />
          </OutlineProvider>
        </LoadingContext.Provider>
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
