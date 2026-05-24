import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import type { IWorkbookData } from '@univerjs/core';
import { TitleBar } from './shell/TitleBar';
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
  type PreviewState,
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
import { VersionHistoryPanel } from './shell/VersionHistoryPanel';
import { PanelRail } from './shell/PanelRail';
import { PanelMutex } from './shell/PanelMutex';
import { PreviewBanner } from './shell/PreviewBanner';
import { PreviewDriver } from './shell/PreviewDriver';
import { InputParserDriver } from './shell/InputParserDriver';
import { ThemeBridge } from './shell/ThemeBridge';
import { HomeScreen } from './home/HomeScreen';
import { ShowFormulasLayer } from './shell/ShowFormulasLayer';
import { PivotsProvider } from './pivots/pivots-context';
import { SparklinesProvider } from './sparklines/sparklines-context';
import { SparklineLayer } from './sparklines/SparklineLayer';
import { useAutosave } from './autosave/useAutosave';
import { AutosaveRestoreBanner } from './autosave/AutosaveRestoreBanner';
import { useVersionHistoryCapture } from './version-history/useVersionHistoryCapture';
import { useTouchPan } from './touch/useTouchPan';
import { MobileActionBar } from './shell/MobileActionBar';

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

  const [homeDismissed, setHomeDismissed] = useState(false);
  // Auto-dismiss the home screen when an autosave record exists, so
  // the AutosaveRestoreBanner is visible on first paint instead of
  // being hidden behind the template gallery. Best-effort IDB probe —
  // failures (private mode, locked DB) leave home visible, which is
  // the conservative default.
  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const { readAutosave } = await import('./autosave/store');
        const rec = await readAutosave();
        if (!cancelled && rec) {
          setHomeDismissed(true);
        }
      } catch {
        /* keep home visible on probe failure */
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);
  const [formulaBarVisible, setFormulaBarVisible] = useState(true);
  const [tablesPanelVisible, setTablesPanelVisible] = useState(false);
  const [outlinePanelVisible, setOutlinePanelVisible] = useState(false);
  const [chartsPanelVisible, setChartsPanelVisible] = useState(false);
  const [historyPanelVisible, setHistoryPanelVisible] = useState(false);
  const [shareRoomOpen, setShareRoomOpen] = useState(false);
  const [showFormulas, setShowFormulas] = useState(false);
  const [loading, setLoading] = useState<LoadingState | null>(null);

  // Version-history preview state. `preview` is the visible shape;
  // `previewSavedRef` keeps the pre-preview workbook off React state
  // so a multi-MB snapshot isn't duplicated when we already have one
  // copy in Univer.
  const [preview, setPreview] = useState<PreviewState | null>(null);
  const previewSavedRef = useRef<{
    data: IWorkbookData;
    sourceFormat: WorkbookFormat | null;
  } | null>(null);

  const enterPreview = useCallback(
    (
      versionId: number,
      versionName: string,
      versionSavedAt: number,
      snapshotData: IWorkbookData,
      snapshotSourceFormat: WorkbookFormat | null,
      currentLiveData: IWorkbookData,
      currentLiveFormat: WorkbookFormat | null,
    ) => {
      previewSavedRef.current = { data: currentLiveData, sourceFormat: currentLiveFormat };
      setPreview({ versionId, versionName, versionSavedAt });
      // Stash both refs before the swap kicks off the React revision
      // bump — the snapshot ref-clear timer must see the new data, not
      // the saved one. `replaceWorkbook` itself sets snapshotRef.
      snapshotRef.current = snapshotData;
      setMeta((prev) => ({
        id: snapshotData.id ?? prev.id,
        name: snapshotData.name ?? prev.name,
        sourceFormat: snapshotSourceFormat,
        revision: prev.revision + 1,
      }));
      // Drop the stash after consumers handle the revision bump.
      setTimeout(() => {
        setTimeout(() => {
          if (snapshotRef.current === snapshotData) snapshotRef.current = null;
        }, 0);
      }, 0);
    },
    [],
  );

  const exitPreview = useCallback(() => {
    const saved = previewSavedRef.current;
    previewSavedRef.current = null;
    setPreview(null);
    if (!saved) return;
    snapshotRef.current = saved.data;
    setMeta((prev) => ({
      id: saved.data.id ?? prev.id,
      name: saved.data.name ?? prev.name,
      sourceFormat: saved.sourceFormat,
      revision: prev.revision + 1,
    }));
    setTimeout(() => {
      setTimeout(() => {
        if (snapshotRef.current === saved.data) snapshotRef.current = null;
      }, 0);
    }, 0);
  }, []);

  const commitPreview = useCallback(() => {
    // The snapshot is already the live workbook (loaded in
    // enterPreview). All we do is drop the saved-state ref and clear
    // the preview flag so editing re-enables and the banner hides.
    previewSavedRef.current = null;
    setPreview(null);
  }, []);

  const replaceWorkbook = useCallback((next: IWorkbookData, format?: WorkbookFormat | null) => {
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
  }, []);

  // App owns only the meta update — TitleBar mirrors into Univer via
  // setName since App itself is outside the UniverProvider.
  const renameWorkbook = useCallback((name: string) => {
    const trimmed = name.trim();
    if (!trimmed) return;
    setMeta((prev) => (prev.name === trimmed ? prev : { ...prev, name: trimmed }));
  }, []);

  const wbValue: WorkbookCtxValue = useMemo(
    () => ({
      meta,
      snapshotRef,
      replaceWorkbook,
      renameWorkbook,
      preview,
      enterPreview,
      exitPreview,
      commitPreview,
    }),
    [meta, replaceWorkbook, renameWorkbook, preview, enterPreview, exitPreview, commitPreview],
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
      // Side panels are mutually exclusive — opening one auto-closes
      // the others. Three of them open at once would squeeze the grid
      // into a strip and competing context (which one am I editing?).
      tablesPanelVisible,
      toggleTablesPanel: () =>
        setTablesPanelVisible((v) => {
          const next = !v;
          if (next) {
            setOutlinePanelVisible(false);
            setChartsPanelVisible(false);
            setHistoryPanelVisible(false);
          }
          return next;
        }),
      outlinePanelVisible,
      toggleOutlinePanel: () =>
        setOutlinePanelVisible((v) => {
          const next = !v;
          if (next) {
            setTablesPanelVisible(false);
            setChartsPanelVisible(false);
            setHistoryPanelVisible(false);
          }
          return next;
        }),
      chartsPanelVisible,
      toggleChartsPanel: () =>
        setChartsPanelVisible((v) => {
          const next = !v;
          if (next) {
            setTablesPanelVisible(false);
            setOutlinePanelVisible(false);
            setHistoryPanelVisible(false);
          }
          return next;
        }),
      historyPanelVisible,
      toggleHistoryPanel: () =>
        setHistoryPanelVisible((v) => {
          const next = !v;
          if (next) {
            setTablesPanelVisible(false);
            setOutlinePanelVisible(false);
            setChartsPanelVisible(false);
          }
          return next;
        }),
      closeAllReactPanels: () => {
        setTablesPanelVisible(false);
        setOutlinePanelVisible(false);
        setChartsPanelVisible(false);
        setHistoryPanelVisible(false);
      },
      showFormulas,
      toggleShowFormulas: () => setShowFormulas((v) => !v),
      openShareRoom: () => setShareRoomOpen(true),
    }),
    [
      formulaBarVisible,
      tablesPanelVisible,
      outlinePanelVisible,
      chartsPanelVisible,
      historyPanelVisible,
      showFormulas,
    ],
  );

  return (
    <UniverRoot>
      <UIContext.Provider value={uiValue}>
        <WorkbookContext.Provider value={wbValue}>
          <LoadingContext.Provider value={loadingValue}>
            <BusyProvider>
              <ChartsProvider>
                <PivotsProvider>
                  <SparklinesProvider>
                    <OutlineProvider>
                      <GrowthDriver />
                      <FileDropDriver />
                      <AutosaveDriver />
                      <TouchPanDriver />
                      <VersionHistoryDriver />
                      <PreviewDriver />
                      <InputParserDriver />
                      <ThemeBridge />
                      <CollabDriver>
                        <div
                          className={`app${formulaBarVisible ? '' : ' app--no-formula-bar'}`}
                          data-testid="app-shell"
                        >
                          <TitleBar />
                          <Toolbar />
                          <AutosaveRestoreBanner />
                          <PreviewBanner />
                          {formulaBarVisible && <FormulaBar />}
                          <div className="grid-row">
                            <main className="grid-host" data-testid="grid-host">
                              <UniverSheet revision={meta.revision} initialSnapshot={initial} />
                            </main>
                            {tablesPanelVisible && <TablesPanel />}
                            {outlinePanelVisible && <OutlinePanel />}
                            {chartsPanelVisible && <ChartsPanel />}
                            {historyPanelVisible && <VersionHistoryPanel />}
                            <PanelRail />
                          </div>
                          <MobileActionBar />
                          <SheetTabs />
                          <PanelMutex />
                          {shareRoomOpen && (
                            <CreateRoomDialog onClose={() => setShareRoomOpen(false)} />
                          )}
                        </div>
                      </CollabDriver>
                      <HomeScreen
                        dismissed={homeDismissed}
                        onDismiss={() => setHomeDismissed(true)}
                      />
                      <LoadingOverlay />
                      <ChartLayer />
                      <SparklineLayer />
                      <ShowFormulasLayer />
                    </OutlineProvider>
                  </SparklinesProvider>
                </PivotsProvider>
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

/** Effect-only — drives the IDB autosave loop. No-op in collab rooms. */
function AutosaveDriver(): ReactNode {
  useAutosave();
  return null;
}

/** Effect-only — translates touch-drag on the Univer canvas into wheel
 *  events so the grid scrolls on mobile. Univer 0.24 doesn't ship native
 *  touch-pan; drop this once it does. */
function TouchPanDriver(): ReactNode {
  useTouchPan();
  return null;
}

/** Effect-only — drives the version-history snapshot capture loop.
 *  Coarse cadence (~10 min while dirty) so it doesn't fight autosave. */
function VersionHistoryDriver(): ReactNode {
  useVersionHistoryCapture();
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
