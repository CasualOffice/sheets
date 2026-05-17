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
import { extendContextMenu } from './context-menu-extensions';
import { registerPlugins } from './univer/plugins';
import { installDevHelpers } from './univer/dev-helpers';
import { timeIt } from './perf';
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

    registerPlugins(univer, hostRef.current);

    timeIt('mount-unit', () => univer.createUnit(UniverInstanceType.UNIVER_SHEET, initialSnapshot));

    // Augment the built-in cell context menu with Merge / Unmerge entries.
    extendContextMenu(univer);

    const api = FUniver.newAPI(univer);
    apiRef.current = api;
    setApi(api);

    const raf = requestAnimationFrame(() => setReady(true));
    const teardownDevHelpers = installDevHelpers(api);

    return () => {
      cancelAnimationFrame(raf);
      // Defer disposal so it can't fire during React's render phase — Univer
      // owns its own React root and a synchronous unmount mid-render warns and
      // leaves the canvas detached.
      const toDispose = univer;
      apiRef.current = null;
      univerRef.current = null;
      setApi(null);
      queueMicrotask(() => toDispose.dispose());
      teardownDevHelpers();
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
    try {
      timeIt('swap-unit', () => {
        if (currentId) disposeUnit?.call(api, currentId);
        createSheet.call(api, snapshot);
      });
      lastRevisionRef.current = revision;
      console.info('[open-xlsx] swap complete');
    } catch (err) {
      console.error('[open-xlsx] swap failed', err);
      throw err;
    }
  }, [revision, ctx]);

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
