import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import type { IWorkbookData, ICommandInfo, IExecutionOptions } from '@univerjs/core';
import { ICommandService } from '@univerjs/core';
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
import { ToastProvider } from './shell/toast/toast-context';
import { ToastContainer } from './shell/toast/ToastContainer';
import { ChartsProvider } from './charts/charts-context';
import { ChartLayer } from './charts/ChartLayer';
import { ChartsPanel } from './shell/ChartsPanel';
import { VersionHistoryPanel } from './shell/VersionHistoryPanel';
import { PanelRail } from './shell/PanelRail';
import { PanelMutex } from './shell/PanelMutex';
import { PreviewBanner } from './shell/PreviewBanner';
import { PreviewDriver } from './shell/PreviewDriver';
import { ThemeBridge } from './shell/ThemeBridge';
import { HomeScreen } from './home/HomeScreen';
import { ShowFormulasLayer } from './shell/ShowFormulasLayer';
import { PivotsProvider } from './pivots/pivots-context';
import { SparklinesProvider } from './sparklines/sparklines-context';
import { SparklineLayer } from './sparklines/SparklineLayer';
import { useAutosave } from './autosave/useAutosave';
import { AutosaveRestoreBanner } from './autosave/AutosaveRestoreBanner';
import { FileSourceProvider, useFileSource } from './file-source';
import { AuthProvider, PersonalAuthGate } from './auth';
import { useAuth } from './auth/auth-context';
import { useVersionHistoryCapture } from './version-history/useVersionHistoryCapture';
import { useTouchPan } from './touch/useTouchPan';
import { MobileActionBar } from './shell/MobileActionBar';
import { navigate, useRoute } from './router';
import { MySpreadsheetsList } from './home/MySpreadsheetsList';
import { useUniverAPI } from './use-univer';

export function App() {
  // Route gate. UX_AUDIT.md §1, §5 — personal-mode IA needs `/home` to
  // render the file picker, not the always-mounted editor. The editor
  // tree still mounts for every other route (sheet / sheet-draft / room
  // / unknown) so workbook open / collab / autosave continue working
  // exactly as before; the only behavioural change is that `/home` no
  // longer shows the empty default workbook with an overlay.
  const route = useRoute();
  const showHomeList = route.kind === 'home' || route.kind === 'templates';

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
    serverFileId: null,
    serverEtag: null,
  }));

  const [homeDismissed, setHomeDismissed] = useState(false);

  // Mirror the open file's name into the browser tab/window title,
  // Office-style ("Book1 — Casual Sheets"). On `/home` or `/templates`
  // keep the original marketing/SEO title (captured once at mount) so
  // bookmarks and shared links stay descriptive. Gating on `route.kind`
  // (not `homeDismissed`) means the title flips the moment the URL
  // changes, including via back/forward — UX_AUDIT.md §2.12.
  const baseTitle = useRef(document.title);
  useEffect(() => {
    document.title = showHomeList
      ? baseTitle.current
      : `${meta.name || 'Untitled'} — Casual Sheets`;
  }, [showHomeList, meta.name]);
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

  const replaceWorkbook = useCallback(
    (
      next: IWorkbookData,
      format?: WorkbookFormat | null,
      server?: { fileId: string | null; etag: string | null } | null,
    ) => {
      snapshotRef.current = next;
      setMeta((prev) => ({
        id: next.id ?? prev.id,
        name: next.name ?? prev.name,
        sourceFormat: format !== undefined ? format : prev.sourceFormat,
        revision: prev.revision + 1,
        // Default to null when the caller didn't supply server info
        // — the new workbook is a fresh load (template / drop / FSA
        // open) without a tracked server identity. Subsequent Save
        // falls into the "new" path.
        serverFileId: server?.fileId ?? null,
        serverEtag: server?.etag ?? null,
        // Reset the user-edit gate. The new workbook starts from a
        // clean slate; first content mutation flips this back true via
        // the EditTracker driver. UX_AUDIT.md §5.
        hasUserEdited: false,
      }));
      // Free the ref after consumers have processed the revision
      // bump. We wait two macrotasks: the first lets React flush its
      // render + useEffect pass (UniverSheet's swap, OutlineProvider's
      // rehydrate); the second is paranoia for any deferred work
      // scheduled inside those effects.
      setTimeout(() => {
        setTimeout(() => {
          if (snapshotRef.current === next) snapshotRef.current = null;
        }, 0);
      }, 0);
    },
    [],
  );

  const updateServerEtag = useCallback((etag: string | null) => {
    setMeta((prev) => (prev.serverEtag === etag ? prev : { ...prev, serverEtag: etag }));
  }, []);

  const markUserEdited = useCallback(() => {
    setMeta((prev) => (prev.hasUserEdited ? prev : { ...prev, hasUserEdited: true }));
  }, []);

  // Cleared at the tail of every successful Save (any source — server,
  // FSA, download). Drives the logout dirty-check (UX_AUDIT.md §2.14):
  // `hasUserEdited === true` after a save means the user typed AFTER
  // the save completed, so logout should still warn. A user who saved
  // and then typed gets warned; one who saved and idled doesn't.
  const markSaved = useCallback(() => {
    setMeta((prev) => (prev.hasUserEdited ? { ...prev, hasUserEdited: false } : prev));
  }, []);

  const updateServerFileId = useCallback((fileId: string | null) => {
    setMeta((prev) => (prev.serverFileId === fileId ? prev : { ...prev, serverFileId: fileId }));
    // Rebind the URL so the draft `/sheet/new` becomes the canonical
    // `/sheet/<id>` after the first successful save. replaceState (not
    // pushState) keeps the history clean — back doesn't return to the
    // ephemeral draft URL. No-op when there's no DOM (SSR / tests).
    if (fileId && typeof window !== 'undefined' && window.location.pathname === '/sheet/new') {
      window.history.replaceState(window.history.state, '', `/sheet/${encodeURIComponent(fileId)}`);
      // Tell router subscribers the URL changed so useRoute() re-reads
      // and the rest of the app sees route.kind flip from sheet-draft
      // → sheet without a re-mount.
      window.dispatchEvent(new CustomEvent('cd:navigate'));
    }
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
      updateServerEtag,
      updateServerFileId,
      markUserEdited,
      markSaved,
      preview,
      enterPreview,
      exitPreview,
      commitPreview,
    }),
    [
      meta,
      replaceWorkbook,
      renameWorkbook,
      updateServerEtag,
      updateServerFileId,
      markUserEdited,
      markSaved,
      preview,
      enterPreview,
      exitPreview,
      commitPreview,
    ],
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
            <AuthProvider>
              <FileSourceProvider>
                <ToastProvider>
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
                            <ThemeBridge />
                            <RouteWorkbookSync replaceWorkbook={replaceWorkbook} />
                            <EditTracker markUserEdited={markUserEdited} />
                            <PersonalAuthGate>
                              <RouteHost
                                routeIsHome={showHomeList}
                                home={<MySpreadsheetsList />}
                                editor={
                                  <>
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
                                            <UniverSheet
                                              revision={meta.revision}
                                              initialSnapshot={initial}
                                            />
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
                                    {/* HomeScreen overlay only on the editor
                                        branch — the dedicated /home view doesn't
                                        need an overlay. */}
                                    <HomeScreen
                                      dismissed={homeDismissed}
                                      onDismiss={() => setHomeDismissed(true)}
                                    />
                                  </>
                                }
                              />
                              <LoadingOverlay />
                              <ChartLayer />
                              <SparklineLayer />
                              <ShowFormulasLayer />
                            </PersonalAuthGate>
                          </OutlineProvider>
                        </SparklinesProvider>
                      </PivotsProvider>
                    </ChartsProvider>
                  </BusyProvider>
                  <ToastContainer />
                </ToastProvider>
              </FileSourceProvider>
            </AuthProvider>
          </LoadingContext.Provider>
        </WorkbookContext.Provider>
      </UIContext.Provider>
    </UniverRoot>
  );
}

/** Auth-aware route switch. Lives inside <PersonalAuthGate> so it can
 *  read useAuth() — the App body itself sits ABOVE <AuthProvider> in
 *  the tree and can't.
 *
 *  Two jobs:
 *  1. Decide between MySpreadsheetsList (the /home file picker) and
 *     the editor body. Personal-mode authenticated + route-is-home →
 *     show the list; otherwise fall through to the editor. This keeps
 *     non-personal deploys (Mode 1 / Mode 2 / Playwright / GitHub
 *     Pages, where `auth.state.kind === 'disabled'`) rendering the
 *     editor at `/` exactly as before.
 *  2. Redirect `/` → `/home` when a personal account session exists.
 *     `parseRoute` reports `/` as kind:'home'; we do the URL canonical-
 *     isation here so refresh / share / bookmark all converge.
 *
 *  UX_AUDIT.md §1, §5. */
function RouteHost({
  routeIsHome,
  home,
  editor,
}: {
  routeIsHome: boolean;
  home: ReactNode;
  editor: ReactNode;
}): ReactNode {
  const auth = useAuth();
  // `loading` counts as personal-active so we don't flash the editor
  // shell during the brief auth-status probe on first paint.
  const personalActive =
    auth.state.kind === 'authenticated' || auth.state.kind === 'loading';
  useEffect(() => {
    if (
      personalActive &&
      typeof window !== 'undefined' &&
      window.location.pathname === '/'
    ) {
      navigate('/home', { replace: true });
    }
  }, [personalActive]);
  return personalActive && routeIsHome ? home : editor;
}

/** Effect-only — listens for the first meaningful content mutation on
 *  the workbook and flips `meta.hasUserEdited`. UX_AUDIT.md §5: the
 *  Save handler uses this to skip create-saves of `/sheet/new` drafts
 *  the user opened but never typed in. Mirrors the noisy-mutation
 *  filter from useAutosave so navigation / focus / selection events
 *  don't accidentally promote a clean draft to a server row. */
function EditTracker({ markUserEdited }: { markUserEdited: () => void }): ReactNode {
  const api = useUniverAPI();
  useEffect(() => {
    if (!api) return;
    // Reach the command service via the facade's private `_injector` —
    // same path useAutosave uses (apps/web/src/autosave/useAutosave.ts).
    // The two hooks share this back door so they stay aligned on what
    // "an edit" means; if Univer renames it, both break together and
    // get fixed together.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const injector = (api as any)._injector as
      | { get: (t: unknown) => unknown }
      | undefined;
    if (!injector) return;
    const cmdSvc = injector.get(ICommandService) as {
      onMutationExecutedForCollab: (
        l: (info: ICommandInfo, options?: IExecutionOptions) => void,
      ) => { dispose: () => void };
    };
    const sub = cmdSvc.onMutationExecutedForCollab((info, options) => {
      if (options?.fromCollab) return; // remote replays don't count
      const id = info?.id ?? '';
      // Same noisy-mutation filter as useAutosave so navigation /
      // selection / sheet-switch don't promote a clean draft.
      if (id.startsWith('sheet.mutation.set-selections')) return;
      if (id === 'sheet.mutation.set-worksheet-active-operation') return;
      markUserEdited();
    });
    return () => sub.dispose();
  }, [api, markUserEdited]);
  return null;
}

/** Effect-only — watches the route and calls `fileSource.openRecent` when
 *  the URL is `/sheet/<id>`. UX_AUDIT.md §5 Phase 1. The default
 *  emptyWorkbook stays in place for `/sheet/new` (the draft route) and
 *  for `/r/<roomId>` (legacy anonymous coedit) where the workbook is
 *  picked up via the collab driver, not the file source. */
function RouteWorkbookSync({
  replaceWorkbook,
}: {
  replaceWorkbook: WorkbookCtxValue['replaceWorkbook'];
}): ReactNode {
  const route = useRoute();
  const fileSource = useFileSource();
  const lastOpenedRef = useRef<string | null>(null);
  useEffect(() => {
    if (route.kind !== 'sheet' || !route.id) return;
    if (lastOpenedRef.current === route.id) return; // already loaded
    let cancelled = false;
    void (async () => {
      try {
        const opened = await fileSource.openRecent(route.id);
        if (cancelled) return;
        replaceWorkbook(
          opened.data,
          opened.sourceFormat,
          opened.serverFileId
            ? { fileId: opened.serverFileId, etag: opened.serverEtag ?? null }
            : null,
        );
        lastOpenedRef.current = route.id;
      } catch (err) {
        if (cancelled) return;
        // Stale URL / file deleted / share token expired — bounce back
        // to the list. Toast goes here once the toast surface is
        // hookable from outside React's render tree.
        // eslint-disable-next-line no-console
        console.warn('[home] could not open sheet', route.id, err);
        navigate('/home');
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [route.kind, route.id, fileSource, replaceWorkbook]);
  return null;
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
