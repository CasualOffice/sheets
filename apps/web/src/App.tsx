import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import type { IWorkbookData } from '@univerjs/core';
import { xlsxToWorkbookData } from './xlsx';
import { odsToWorkbookData, csvToWorkbookData, tsvToWorkbookData } from './ods';
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
import { BusyProvider } from './busy-context';
import { ChartsProvider } from './charts/charts-context';
import { ChartLayer } from './charts/ChartLayer';
import { ChartsPanel } from './shell/ChartsPanel';

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
  const [chartsPanelVisible, setChartsPanelVisible] = useState(false);
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
      // Inside Casual Office (desktop, single-user, no memory pressure),
      // skip the auto-clear. React 18 concurrent rendering can defer
      // UniverSheet's swap effect past the 2-macrotask setTimeout chain,
      // leaving the swap to find an empty ref — visible as a permanently
      // blank canvas with "swap aborted: snapshotRef is empty" in console.
      // The web build keeps the original aggressive GC behavior.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      if (typeof window !== 'undefined' && (window as any).__deskApp__?.isDesktop) {
        return;
      }
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

  // When loaded inside the Casual Office Tauri shell, the desk-bridge
  // bootstrap defines window.__deskApp__ with the file path the user
  // opened. Read it through the bridge and replace the empty workbook.
  useEffect(() => {
    const bridge = typeof window !== 'undefined' ? window.__deskApp__ : undefined;
    if (!bridge?.isDesktop || !bridge.filePath) return;
    let cancelled = false;
    (async () => {
      const path = bridge.filePath!;
      const fileName = path.split(/[\\/]/).pop() || 'Workbook.xlsx';
      const lower = fileName.toLowerCase();
      let format: WorkbookFormat = 'xlsx';
      if (lower.endsWith('.ods')) format = 'ods';
      else if (lower.endsWith('.csv')) format = 'csv';
      else if (lower.endsWith('.tsv') || lower.endsWith('.tab')) format = 'tsv';
      try {
        setLoading({ fileName, phase: 'reading' });
        const buffer = await bridge.loadDocument();
        if (cancelled) return;
        setLoading({ fileName, phase: 'parsing' });
        let data: IWorkbookData;
        if (format === 'ods') data = await odsToWorkbookData(buffer);
        else if (format === 'csv') data = await csvToWorkbookData(buffer);
        else if (format === 'tsv') data = await tsvToWorkbookData(buffer);
        else data = await xlsxToWorkbookData(buffer);
        if (cancelled) return;
        data.name = fileName.replace(/\.(xlsx|xlsm|ods|csv|tsv|tab)$/i, '');
        setLoading({ fileName, phase: 'mounting' });
        replaceWorkbook(data, format);
        setLoading(null);
      } catch (err) {
        console.error('deskApp load failed', err);
        if (!cancelled) {
          setLoading({ fileName, phase: 'reading', error: String(err) });
        }
      }
    })();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

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
              onRetry: next.onRetry,
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
      chartsPanelVisible,
      toggleChartsPanel: () => setChartsPanelVisible((v) => !v),
      openShareRoom: () => setShareRoomOpen(true),
    }),
    [formulaBarVisible, tablesPanelVisible, outlinePanelVisible, chartsPanelVisible],
  );

  return (
    <UniverRoot>
      <UIContext.Provider value={uiValue}>
        <WorkbookContext.Provider value={wbValue}>
        <LoadingContext.Provider value={loadingValue}>
        <BusyProvider>
        <ChartsProvider>
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
                  {chartsPanelVisible && <ChartsPanel />}
                </div>
                <SheetTabs />
                {shareRoomOpen && (
                  <CreateRoomDialog onClose={() => setShareRoomOpen(false)} />
                )}
              </div>
            </CollabDriver>
            <LoadingOverlay />
            <ChartLayer />
          </OutlineProvider>
        </ChartsProvider>
        </BusyProvider>
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
