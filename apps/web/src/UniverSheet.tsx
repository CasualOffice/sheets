import { useContext, useEffect, useRef, useState } from 'react';
import type { IWorkbookData, Univer } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import { CasualSheets, type CasualSheetsAPI } from '@casualoffice/sheets/sheets';

// Side-effect modules — must load before Univer is constructed (inside the SDK).
// `./univer/styles` pulls in every Univer plugin's CSS (incl. crosshair + zen
// editor, which the SDK doesn't bundle); `./univer/facade` adds the lazy-plugin
// FUniver mixins (cf / dv / hyperlink / …) the app's shell calls.
import './univer/styles';
import './univer/facade';

import { LOCALES } from './locale';
import { useSetUniverAPI } from './use-univer';
import { WorkbookContext } from './workbook-context';
import { registerExtraPlugins } from './univer/extra-plugins';
import { extendContextMenu } from './context-menu-extensions';
import { registerPasteMergeHook } from './univer/paste-merge-hook';
import { installDevHelpers } from './univer/dev-helpers';
import { disableUniverZoomShortcut } from './univer/disable-zoom-shortcut';

type Props = {
  /** First snapshot used to mount the editor. Only consulted on initial mount;
   *  subsequent swaps come through `WorkbookContext.snapshotRef` + `revision`. */
  initialSnapshot: IWorkbookData;
  /** Bumped by `replaceWorkbook` — drives the swap effect. */
  revision: number;
};

/**
 * The app's editor mount. Phase 3: this is now a THIN host over the SDK editor
 * core — it renders `<CasualSheets chrome="none">` (which owns the Univer
 * bootstrap, plugin set, formula engine, snapshot/API) and layers the app's
 * extra concerns on top:
 *   - `onBeforeCreateUnit` registers the EXTRA plugins the SDK doesn't bundle
 *     (crosshair highlight, zen editor) + the Merge/Unmerge context-menu items.
 *   - `formula={{ worker }}` moves compute off-thread (the app's standard
 *     formula worker).
 *   - `onReady(api)` publishes the FUniver facade to `WorkbookContext` and wires
 *     the paste-merge hook, dev helpers, and zoom-shortcut override.
 *   - the revision effect swaps the workbook via `api.loadSnapshot`.
 * The rich Office shell (apps/web/src/shell/) mounts around this, unchanged.
 */
export function UniverSheet({ initialSnapshot, revision }: Props) {
  const setApi = useSetUniverAPI();
  const ctx = useContext(WorkbookContext);
  const [ready, setReady] = useState(false);

  const apiRef = useRef<CasualSheetsAPI | null>(null);
  const teardownsRef = useRef<Array<() => void>>([]);

  // One off-main formula worker for the editor's lifetime. Created lazily on
  // first render (a ref, so React strict-mode's double render reuses it) and
  // terminated on unmount.
  const workerRef = useRef<Worker | null>(null);
  if (!workerRef.current) {
    workerRef.current = new Worker(new URL('./univer/formula-worker.ts', import.meta.url), {
      type: 'module',
      name: 'formula-worker',
    });
  }

  // Register-time extras — run with the raw Univer BEFORE the unit is created.
  const onBeforeCreateUnit = (univer: Univer) => {
    registerExtraPlugins(univer);
    // Merge / Unmerge context-menu entries (uses univer.__getInjector()).
    extendContextMenu(univer);
  };

  // Post-createUnit wiring — all these take the FUniver facade (`api.univer`).
  const handleReady = (api: CasualSheetsAPI) => {
    apiRef.current = api;
    // CRITICAL: publish the FUniver facade (not the CasualSheetsAPI wrapper) —
    // useUniverAPI() is typed FUniver and the whole shell calls facade methods.
    setApi(api.univer);

    const teardowns: Array<() => void> = [];
    const pasteMerge = registerPasteMergeHook(api.univer);
    if (pasteMerge) teardowns.push(pasteMerge);
    teardowns.push(installDevHelpers(api.univer));
    teardowns.push(disableUniverZoomShortcut(api.univer));
    teardownsRef.current = teardowns;

    // Force a workbook-wide recalc so formula cells shipped without cached <v>
    // (hand-authored templates) populate on first paint.
    runInitialRecalc(api.univer);

    requestAnimationFrame(() => setReady(true));
  };

  // Teardown on unmount: dispose the layered hooks, clear the context, kill the
  // worker. (CasualSheets disposes the Univer instance itself.)
  useEffect(
    () => () => {
      for (const t of teardownsRef.current) t();
      teardownsRef.current = [];
      apiRef.current = null;
      setApi(null);
      workerRef.current?.terminate();
      workerRef.current = null;
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [],
  );

  // Swap the workbook when `revision` bumps (File → Open, collab snapshot
  // replace). Serialised through swapChainRef so back-to-back swaps (collab
  // seed → compaction with the same unit id) don't race into a duplicate-unit-id
  // throw. The next snapshot rides on ctx.snapshotRef for the brief window
  // between replaceWorkbook and this effect.
  const lastRevisionRef = useRef<number>(revision);
  const swapChainRef = useRef<Promise<void>>(Promise.resolve());
  useEffect(() => {
    if (lastRevisionRef.current === revision) return;
    lastRevisionRef.current = revision;
    const api = apiRef.current;
    const snapshot = ctx?.snapshotRef.current ?? null;
    if (!api) {
      console.warn('[open-xlsx] swap aborted: api not ready yet');
      return;
    }
    if (!snapshot) {
      console.warn('[open-xlsx] swap aborted: snapshotRef empty for revision', revision);
      return;
    }
    swapChainRef.current = swapChainRef.current
      .catch(() => undefined)
      .then(async () => {
        try {
          // loadSnapshot disposes the current unit and mounts the new one. The
          // SDK has already idle-loaded the feature plugins by now, so the new
          // snapshot's CF / table / hyperlink resources resolve.
          await api.loadSnapshot(snapshot);
          runInitialRecalc(api.univer);
          console.info('[open-xlsx] swap complete', { to: snapshot.id });
        } catch (err) {
          console.error('[open-xlsx] swap failed', err);
          throw err;
        }
      });
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
      <CasualSheets
        initialData={initialSnapshot}
        locales={LOCALES}
        chrome="none"
        testId="univer-host"
        formula={{ worker: workerRef.current }}
        onBeforeCreateUnit={onBeforeCreateUnit}
        onReady={handleReady}
        style={{ width: '100%', height: '100%' }}
      />
    </>
  );
}

/**
 * Trigger a one-shot workbook-wide formula recalc after mount/swap. Templates
 * ship as `.xlsx`; hand-authored ones often omit the cached `<v>` next to
 * `<f>`, so formulas land with `f` set and `v` empty — blank until something
 * forces compute. `getFormula().executeCalculation()` recomputes everything.
 * The formula facade loads lazily, so retry once on a microtask if absent.
 */
function runInitialRecalc(api: FUniver): void {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const formula = (api as any).getFormula?.() as { executeCalculation: () => void } | undefined;
    if (formula?.executeCalculation) {
      formula.executeCalculation();
      return;
    }
  } catch (err) {
    console.warn('[recalc] initial recalc failed', err);
  }
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
