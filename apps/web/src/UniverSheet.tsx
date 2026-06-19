import { useContext, useEffect, useRef, useState } from 'react';
import {
  LocaleType,
  LogLevel,
  Univer,
  UniverInstanceType,
  type IWorkbookData,
} from '@univerjs/core';
import { FUniver } from '@univerjs/core/facade';
import type { FWorkbook } from '@univerjs/sheets/facade';
import { defaultTheme } from '@univerjs/themes';

// Side-effect modules — must load before Univer is constructed.
import './univer/styles';
import './univer/facade';

import { LOCALES } from './locale';
import { useSetUniverAPI } from './use-univer';
import { useLoading } from './loading-context';
import { extendContextMenu } from './context-menu-extensions';
import { registerPlugins } from './univer/plugins';
import { eagerLoadForSnapshot, idleLoadAll, setUniverForLazyLoad } from '@casualoffice/sheets/univer';
import { installDevHelpers } from './univer/dev-helpers';
import { disableUniverZoomShortcut } from './univer/disable-zoom-shortcut';
import { registerPasteMergeHook } from './univer/paste-merge-hook';
import { timeIt, timeItAsync } from './perf';
import { WorkbookContext } from './workbook-context';

type Props = {
  /** First snapshot used to mount Univer. Only consulted on initial mount;
   *  subsequent swaps come through `WorkbookContext.snapshotRef`. */
  initialSnapshot: IWorkbookData;
  /** Bumped by `replaceWorkbook` — drives the swap effect without us having
   *  to hold a snapshot reference in React state (Stage 3 memory win). */
  revision: number;
};

export function UniverSheet({ initialSnapshot, revision }: Props) {
  const hostRef = useRef<HTMLDivElement>(null);
  const [ready, setReady] = useState(false);
  const setApi = useSetUniverAPI();
  const ctx = useContext(WorkbookContext);
  const loading = useLoading();
  // Hold the live Univer instance + API across snapshot swaps so Open replaces
  // the workbook unit rather than tearing the whole Univer (and its internal
  // React root) down — the latter races React's render phase and leaves the
  // grid blank.
  const univerRef = useRef<Univer | null>(null);
  const apiRef = useRef<FUniver | null>(null);

  useEffect(() => {
    if (!hostRef.current) return;

    const univer = new Univer({
      theme: defaultTheme,
      locale: LocaleType.EN_US,
      locales: LOCALES,
      logLevel: LogLevel.WARN,
    });
    univerRef.current = univer;
    setUniverForLazyLoad(univer);

    registerPlugins(univer, hostRef.current);

    let teardownDevHelpers: (() => void) | null = null;
    let teardownPasteMergeHook: (() => void) | null = null;
    let teardownZoomShortcut: (() => void) | null = null;
    let raf = 0;
    let cancelled = false;
    void (async () => {
      // Eager-load any plugin the initial snapshot needs (CF rules,
      // table defs, hyperlinks, etc.) BEFORE createUnit. Skipping this
      // would silently drop the resource keys for plugins that aren't
      // registered when Univer's resource manager reads the snapshot.
      await timeItAsync('eager-plugins', () => eagerLoadForSnapshot(univer, initialSnapshot));
      if (cancelled) return;
      timeIt('mount-unit', () =>
        univer.createUnit(UniverInstanceType.UNIVER_SHEET, initialSnapshot),
      );

      // Augment the built-in cell context menu with Merge / Unmerge entries.
      extendContextMenu(univer);

      const api = FUniver.newAPI(univer);
      apiRef.current = api;
      setApi(api);

      // Force a workbook-wide recalc so formula cells that ship without
      // cached `<v>` values (common in hand-authored templates) populate
      // on first paint — otherwise they render blank until the user
      // edits a cell and triggers cascade. Wrapped in try/catch because
      // the formula facade is loaded lazily and might not be ready on
      // very fresh installs; the worst case is a deferred recalc.
      runInitialRecalc(api);

      // Excel-paste merge preservation: registers a clipboard hook
      // that emits AddWorksheetMergeMutation for any colspan/rowspan
      // cells in the pasted HTML. Without this, pasting a merged
      // range from Excel lands as N separate cells.
      teardownPasteMergeHook = registerPasteMergeHook(api);

      raf = requestAnimationFrame(() => setReady(true));
      teardownDevHelpers = installDevHelpers(api);
      // Override Univer's Ctrl+- / Ctrl+= zoom shortcuts so our
      // Excel-style Insert/Delete-cells dialogs aren't fighting a
      // simultaneous canvas zoom.
      teardownZoomShortcut = disableUniverZoomShortcut(api);

      // Idle-load every remaining lazy plugin so the user finds them
      // ready when they reach the Insert / Data / Format tabs. The
      // bundle split (per Stage 4) is the persistent boot-time win;
      // this just ensures runtime feature parity with the previous
      // monolithic mount.
      idleLoadAll(univer);
    })();

    return () => {
      cancelled = true;
      cancelAnimationFrame(raf);
      // Defer disposal so it can't fire during React's render phase — Univer
      // owns its own React root and a synchronous unmount mid-render warns and
      // leaves the canvas detached.
      const toDispose = univer;
      apiRef.current = null;
      univerRef.current = null;
      setUniverForLazyLoad(null);
      setApi(null);
      queueMicrotask(() => toDispose.dispose());
      teardownDevHelpers?.();
      teardownPasteMergeHook?.();
      teardownZoomShortcut?.();
    };
    // Mount Univer exactly once. Snapshot changes are handled by the swap
    // effect below — recreating Univer per snapshot would race React's render.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Swap the active workbook unit when `revision` bumps. We never read
  // the next snapshot from React state — it lives on `ctx.snapshotRef`
  // for the brief window between replaceWorkbook and this effect, then
  // gets cleared so the workbook tree is GC-eligible.
  //
  // CRITICAL: swaps are serialised through `swapChainRef`. Without that,
  // back-to-back replaceWorkbook calls (e.g. owner loads seed → bridge
  // immediately replays a compaction snapshot with the same workbook
  // id) start two concurrent async swaps. Both await eager-plugins,
  // both read `current` before either dispose runs, then both call
  // createUnit with the same id — Univer throws
  // "cannot create a unit with the same unit id: wb-...".
  const lastRevisionRef = useRef<number>(revision);
  const swapChainRef = useRef<Promise<void>>(Promise.resolve());
  useEffect(() => {
    if (lastRevisionRef.current === revision) return;
    lastRevisionRef.current = revision; // claim this revision immediately so retries don't re-fire
    const api = apiRef.current;
    if (!api) {
      console.warn('[open-xlsx] swap aborted: api not ready yet');
      return;
    }
    const snapshot = ctx?.snapshotRef.current ?? null;
    if (!snapshot) {
      console.warn('[open-xlsx] swap aborted: snapshotRef is empty for revision', revision);
      return;
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const apiAny = api as any;
    const createSheet = apiAny.createUniverSheet as ((data: IWorkbookData) => unknown) | undefined;
    const disposeUnit = apiAny.disposeUnit as ((id: string) => void) | undefined;
    if (!createSheet) {
      console.warn('[open-xlsx] swap aborted: createUniverSheet missing on facade');
      return;
    }
    // Chain onto the in-flight swap so two revisions back-to-back run
    // sequentially. .catch swallows the previous error to keep the
    // chain alive — the previous swap already reported via the loading
    // overlay.
    swapChainRef.current = swapChainRef.current
      .catch(() => undefined)
      .then(async () => {
        try {
          const u = univerRef.current;
          if (u) {
            await timeItAsync('eager-plugins', () => eagerLoadForSnapshot(u, snapshot));
          }
          timeIt('swap-unit', () => {
            // Re-read `current` HERE, AFTER eager-load + after any
            // previous chained swap completed. Reading it before the
            // chain wait would give us a stale unit id that the previous
            // swap already disposed.
            const current = api.getActiveWorkbook() as unknown as FWorkbook | null;
            const currentId = current?.getId();
            // Defensive: if a unit with the snapshot's id already exists
            // (e.g. someone called replaceWorkbook twice with the same
            // data), dispose it first so createUnit doesn't collide.
            if (currentId && currentId !== snapshot.id) {
              disposeUnit?.call(api, currentId);
            }
            if (snapshot.id && currentId !== snapshot.id) {
              // Dispose any orphaned unit holding the target id.
              try {
                disposeUnit?.call(api, snapshot.id);
              } catch {
                /* fine — unit didn't exist */
              }
            } else if (currentId === snapshot.id) {
              // Same id as the current unit — dispose it explicitly so
              // createUnit gets a clean slot.
              disposeUnit?.call(api, currentId);
            }
            createSheet.call(api, snapshot);
          });
          runInitialRecalc(api);
          console.info('[open-xlsx] swap complete', { to: snapshot.id });
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          console.error('[open-xlsx] swap failed', err);
          loading.set({
            fileName: snapshot.name ?? 'workbook',
            phase: 'mounting',
            error: `Couldn't mount the workbook: ${msg}`,
          });
          throw err; // keep the chain's error visible to the next .catch
        }
      });
  }, [revision, ctx, loading]);

  return (
    <>
      {!ready && (
        <div className="grid-skeleton" data-testid="grid-skeleton" aria-hidden="true">
          <span className="grid-skeleton__chip">
            <span className="grid-skeleton__spinner" />
            Loading workbook…
          </span>
        </div>
      )}
      <div ref={hostRef} data-testid="univer-host" />
    </>
  );
}

/**
 * Trigger a one-shot workbook-wide formula recalc immediately after
 * mount/swap. Templates ship as `.xlsx` files; ExcelJS only populates
 * a cell's `result` (which becomes Univer's `v`) when the OOXML has a
 * cached `<v>` next to the `<f>`. Hand-authored templates often omit
 * the cache, so formulas land in Univer with `f` set and `v` empty —
 * the renderer shows a blank cell until something forces the engine
 * to compute. Editing any cell triggers a cascade and reveals all
 * pending formulas, which is exactly the "press Enter once and it
 * appears" workaround the user reported.
 *
 * `getFormula().executeCalculation()` dispatches the trigger mutation
 * with `forceCalculation: true`, which the calc controller reads as
 * "recompute everything regardless of dirty state."
 */
function runInitialRecalc(api: FUniver): void {
  try {
    // The facade extension is loaded lazily; if it hasn't been
    // initialised yet, retry on the next microtask. Logging on the
    // outer call site would be noisy — the recalc is idempotent and
    // safe to retry.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const formula = (api as any).getFormula?.() as { executeCalculation: () => void } | undefined;
    if (formula?.executeCalculation) {
      formula.executeCalculation();
      return;
    }
  } catch (err) {
    console.warn('[recalc] initial recalc failed', err);
  }
  // Best-effort retry: the formula plugin loads on a microtask, so a
  // queueMicrotask defer gives it a chance to register before we give
  // up entirely.
  queueMicrotask(() => {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const formula = (api as any).getFormula?.() as { executeCalculation: () => void } | undefined;
      formula?.executeCalculation?.();
    } catch {
      /* give up — user can force recalc with F9 */
    }
  });
}
