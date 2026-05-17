import { useEffect, useRef, useState } from 'react';
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

type Props = { snapshot: IWorkbookData };

export function UniverSheet({ snapshot }: Props) {
  const hostRef = useRef<HTMLDivElement>(null);
  const [ready, setReady] = useState(false);
  const setApi = useSetUniverAPI();
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

    timeIt('mount-unit', () => univer.createUnit(UniverInstanceType.UNIVER_SHEET, snapshot));

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

  // Swap the active workbook unit when the snapshot changes, without
  // tearing down the Univer instance itself. Dispose the old unit BEFORE
  // creating the new one so we never collide on unit ids (Univer's
  // IUniverInstanceService throws on duplicate ids).
  const lastSnapshotRef = useRef<IWorkbookData>(snapshot);
  useEffect(() => {
    if (lastSnapshotRef.current === snapshot) return;
    const api = apiRef.current;
    if (!api) {
      console.warn('[open-xlsx] swap aborted: api not ready yet');
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
      lastSnapshotRef.current = snapshot;
      console.info('[open-xlsx] swap complete');
    } catch (err) {
      console.error('[open-xlsx] swap failed', err);
      throw err;
    }
  }, [snapshot]);

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
