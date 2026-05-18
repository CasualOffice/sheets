import { useContext, useEffect, useRef, useState } from 'react';
import { LocaleType, LogLevel, Univer, UniverInstanceType, type IWorkbookData } from '@univerjs/core';
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
import {
  eagerLoadForSnapshot,
  idleLoadAll,
  setUniverForLazyLoad,
} from './univer/lazy-plugins';
import { installDevHelpers } from './univer/dev-helpers';
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
    let raf = 0;
    let cancelled = false;
    void (async () => {
      // Eager-load any plugin the initial snapshot needs (CF rules,
      // table defs, hyperlinks, etc.) BEFORE createUnit. Skipping this
      // would silently drop the resource keys for plugins that aren't
      // registered when Univer's resource manager reads the snapshot.
      await timeItAsync('eager-plugins', () => eagerLoadForSnapshot(univer, initialSnapshot));
      if (cancelled) return;
      timeIt('mount-unit', () => univer.createUnit(UniverInstanceType.UNIVER_SHEET, initialSnapshot));

      // Augment the built-in cell context menu with Merge / Unmerge entries.
      extendContextMenu(univer);

      const api = FUniver.newAPI(univer);
      apiRef.current = api;
      setApi(api);

      // Excel-paste merge preservation: registers a clipboard hook
      // that emits AddWorksheetMergeMutation for any colspan/rowspan
      // cells in the pasted HTML. Without this, pasting a merged
      // range from Excel lands as N separate cells.
      teardownPasteMergeHook = registerPasteMergeHook(api);

      raf = requestAnimationFrame(() => setReady(true));
      teardownDevHelpers = installDevHelpers(api);

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
    };
    // Mount Univer exactly once. Snapshot changes are handled by the swap
    // effect below — recreating Univer per snapshot would race React's render.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Swap the active workbook unit when `revision` bumps. We never read
  // the next snapshot from React state — it lives on `ctx.snapshotRef`
  // for the brief window between replaceWorkbook and this effect, then
  // gets cleared so the workbook tree is GC-eligible.
  const lastRevisionRef = useRef<number>(revision);
  useEffect(() => {
    if (lastRevisionRef.current === revision) return;
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
    const current = api.getActiveWorkbook() as unknown as FWorkbook | null;
    const currentId = current?.getId();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const apiAny = api as any;
    const createSheet = apiAny.createUniverSheet as
      | ((data: IWorkbookData) => unknown)
      | undefined;
    const disposeUnit = apiAny.disposeUnit as ((id: string) => void) | undefined;
    if (!createSheet) {
      console.warn('[open-xlsx] swap aborted: createUniverSheet missing on facade');
      return;
    }
    console.info('[open-xlsx] swapping unit', { from: currentId, to: snapshot.id });
    // Eager-load any plugin the new snapshot needs BEFORE swap, then
    // run the swap synchronously inside a timed block. The await is
    // unavoidable for fresh feature plugins — but most opens hit
    // ones we've already loaded (cached `loaded` set), so this is
    // effectively a no-op fast path after warm-up.
    void (async () => {
      try {
        const u = univerRef.current;
        if (u) {
          await timeItAsync('eager-plugins', () => eagerLoadForSnapshot(u, snapshot));
        }
        timeIt('swap-unit', () => {
          if (currentId) disposeUnit?.call(api, currentId);
          createSheet.call(api, snapshot);
        });
        lastRevisionRef.current = revision;
        console.info('[open-xlsx] swap complete');
      } catch (err) {
        // Surface to the loading overlay so a failed swap doesn't leave
        // the user staring at a forever-spinner. The overlay's error
        // mode stays open until the user dismisses it.
        const msg = err instanceof Error ? err.message : String(err);
        console.error('[open-xlsx] swap failed', err);
        loading.set({
          fileName: snapshot.name ?? 'workbook',
          phase: 'mounting',
          error: `Couldn't mount the workbook: ${msg}`,
        });
      }
    })();
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
