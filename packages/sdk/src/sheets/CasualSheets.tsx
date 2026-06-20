/**
 * CasualSheets — minimal React wrapper around Univer Sheets.
 *
 * Boots Univer with the eager plugin set (render + formula engine +
 * UI + docs + sheets + sheets-ui + sheets-formula + numfmt), mounts a
 * single workbook unit from `initialData`, and hands the host the
 * `CasualSheetsAPI` imperative ref via `onReady` (raw FUniver facade
 * available at `api.univer`).
 *
 * Feature plugins (conditional formatting, data validation, drawings,
 * sort, filter, hyperlinks, tables, comments, find/replace) load lazily
 * by default (`lazyPlugins`): eagerly before mount for whatever the
 * snapshot already uses, idle-loaded otherwise. Pass `lazyPlugins={false}`
 * for the minimal editor.
 *
 * Intentionally NOT included (host can layer on top via FUniver):
 *   - Formula compute via Web Worker — `notExecuteFormula: false`
 *     is the default; the formula engine runs on the main thread.
 *     Host wires `UniverRPCMainThreadPlugin` + a worker URL itself
 *     if it wants the off-main path.
 *   - Snapshot swap (this component mounts a snapshot once; change
 *     the React `key` to remount with a fresh snapshot).
 *   - Paste-merge hooks, dev helpers, zoom-shortcut overrides,
 *     facade extensions — all app concerns.
 *
 * Styles: host must import `@casualoffice/sheets/styles.css`
 * (or the per-plugin CSS) once at app boot. Tree-shaking strips the
 * styles from this entry if the host doesn't reach the styles export.
 */

import { useEffect, useRef, useState, type CSSProperties } from 'react';
import {
  ICommandService,
  LocaleType,
  LogLevel,
  ThemeService,
  Univer,
  UniverInstanceType,
  type ICommandInfo,
  type IExecutionOptions,
  type IWorkbookData,
  type ILocales,
} from '@univerjs/core';
import { FUniver } from '@univerjs/core/facade';
import { defaultTheme } from '@univerjs/themes';

import { UniverRenderEnginePlugin } from '@univerjs/engine-render';
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula';
import { UniverUIPlugin } from '@univerjs/ui';
import { UniverDocsPlugin } from '@univerjs/docs';
import { UniverDocsUIPlugin } from '@univerjs/docs-ui';
import { UniverSheetsPlugin } from '@univerjs/sheets';
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui';
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula';
import { UniverSheetsFormulaUIPlugin } from '@univerjs/sheets-formula-ui';
import { UniverSheetsNumfmtPlugin } from '@univerjs/sheets-numfmt';
import { UniverSheetsNumfmtUIPlugin } from '@univerjs/sheets-numfmt-ui';

import { createCasualSheetsAPI, type CasualSheetsAPI } from './api';
import { eagerLoadForSnapshot, idleLoadAll, setUniverForLazyLoad } from '../univer/lazy-plugins';
import { Toolbar, FormulaBar, StatusBar } from '../chrome';

export interface CasualSheetsProps {
  /** Workbook snapshot to mount. Read once on initial mount; change
   *  the React `key` on this component to remount with a new
   *  workbook. */
  initialData: IWorkbookData;
  /** Called after the workbook unit is created. Hands back the
   *  `CasualSheetsAPI` imperative ref — the SDK's stable integration
   *  surface (snapshot I/O, xlsx import, selection, command dispatch).
   *  The raw FUniver facade is on `api.univer` as the escape hatch. */
  onReady?: (api: CasualSheetsAPI) => void;
  /** Debounced stream of workbook snapshots, emitted after edits
   *  settle. This is the "host persists it" half of the Excalidraw
   *  model — the editor stays storage-unaware and the host writes the
   *  snapshot wherever it likes (localStorage, server, …). Driven by
   *  Univer's mutation hook (`onMutationExecutedForCollab`), not UI
   *  events, so it captures every edit including programmatic ones.
   *  May fire for background/structural mutations too; treat each call
   *  as "current state, persist if you care". */
  onChange?: (snapshot: IWorkbookData) => void;
  /** Debounce window for `onChange`, in ms. Default 400. */
  onChangeDebounceMs?: number;
  /** Lazy-load the feature plugins (conditional formatting, data
   *  validation, hyperlinks, notes, tables, comments, drawings, sort,
   *  filter, find/replace). Default `true`: plugins whose data is in
   *  `initialData` load eagerly before mount (so nothing is dropped on
   *  open), the rest idle-load after first paint. Set `false` for the
   *  minimal editor (render + formula + numfmt only) — the embed-iframe
   *  build does this to stay a single self-contained bundle. */
  lazyPlugins?: boolean;
  /** Locale identifier. Defaults to `LocaleType.EN_US`. */
  locale?: LocaleType;
  /** Locale string bundle. Optional — Univer's default English
   *  strings load if omitted. */
  locales?: ILocales;
  /** Univer log level. Defaults to `LogLevel.WARN`. */
  logLevel?: LogLevel;
  /** Univer chrome toggles. Defaults: header / toolbar / footer off,
   *  context menu on — matches Casual Sheets' embedded shape. */
  ui?: {
    header?: boolean;
    toolbar?: boolean;
    footer?: boolean;
    contextMenu?: boolean;
  };
  /** Override the Univer theme object (colour palette). Defaults to
   *  Univer's `defaultTheme`. Distinct from `appearance` (light/dark). */
  theme?: typeof defaultTheme;
  /** Light or dark mode. Reactive — flipping it re-themes the live
   *  editor via `ThemeService.setDarkMode` (canvas colours, notifications,
   *  and Univer's own `univer-dark` class). Defaults to light.
   *  Note: Univer's Workbench applies the `univer-dark` class to the
   *  document root (`<html>`) itself, so dark mode is page-global by
   *  Univer's design — a host that embeds the editor inside a light page
   *  should scope the editor or accept the global dark CSS. */
  appearance?: 'light' | 'dark';
  /** Office chrome level rendered around the grid:
   *  - `'none'` (default): bare grid — the host supplies its own chrome.
   *  - `'minimal'` / `'full'`: a built-in toolbar (undo/redo/bold/italic/
   *    underline). The rich Office shell (formula bar, menus, status bar) is
   *    being lifted from the app behind `'full'`; until then both render the
   *    minimal toolbar. See SDK_MIGRATION_PIPELINE Phase 1 step 2. */
  chrome?: 'none' | 'minimal' | 'full';
  /** Container style. Default fills the parent. */
  style?: CSSProperties;
  /** Container className for additional styling hooks. */
  className?: string;
  /** Optional test id for the host container. */
  testId?: string;
}

const DEFAULT_STYLE: CSSProperties = {
  width: '100%',
  height: '100%',
  position: 'relative',
};

const DEFAULT_UI = {
  header: false,
  toolbar: false,
  footer: false,
  contextMenu: true,
};

export function CasualSheets({
  initialData,
  onReady,
  onChange,
  onChangeDebounceMs = 400,
  lazyPlugins = true,
  locale = LocaleType.EN_US,
  locales,
  logLevel = LogLevel.WARN,
  ui,
  theme = defaultTheme,
  appearance = 'light',
  chrome = 'none',
  style,
  className,
  testId = 'casual-sheets',
}: CasualSheetsProps) {
  const hostRef = useRef<HTMLDivElement>(null);
  // Keep the latest onChange callable without re-subscribing (the effect
  // mounts once). The subscription itself is only wired when onChange was
  // present at mount.
  const onChangeRef = useRef(onChange);
  onChangeRef.current = onChange;
  const hasOnChange = useRef(!!onChange).current;
  // The live FUniver facade, captured at mount so the reactive appearance
  // effect can reach Univer's ThemeService without re-running boot.
  const apiRef = useRef<CasualSheetsAPI | null>(null);
  // The live API as state, so the built-in chrome (FormulaBar) re-renders and
  // subscribes once the editor is ready. Only set when chrome is shown — the
  // bare-grid path never triggers this re-render. A single post-mount setState
  // doesn't disturb the grid (Univer owns its canvas outside React).
  const [chromeApi, setChromeApi] = useState<CasualSheetsAPI | null>(null);

  useEffect(() => {
    const container = hostRef.current;
    if (!container) return;

    const univer = new Univer({
      theme,
      locale,
      locales,
      logLevel,
    });

    const uiOpts = { ...DEFAULT_UI, ...ui, container };

    univer.registerPlugin(UniverRenderEnginePlugin);
    // Formula compute runs on the MAIN THREAD (no RPC worker). The library build
    // externalises @univerjs (see tsup.config.ts), so the host provides a single
    // redi/@univerjs copy — the duplicate-redi that previously made the formula
    // plugins fail (and which the SDK worked around by dropping them) is gone.
    // apps/web offloads compute to a worker for perf on huge sheets, but a worker
    // shipped inside a published package is brittle to bundle in arbitrary hosts;
    // main-thread keeps the SDK editor self-contained. SDK restructure Batch 2.
    univer.registerPlugin(UniverFormulaEnginePlugin);
    univer.registerPlugin(UniverUIPlugin, uiOpts);
    univer.registerPlugin(UniverDocsPlugin);
    univer.registerPlugin(UniverDocsUIPlugin);
    univer.registerPlugin(UniverSheetsPlugin);
    univer.registerPlugin(UniverSheetsUIPlugin);
    univer.registerPlugin(UniverSheetsFormulaPlugin);
    univer.registerPlugin(UniverSheetsFormulaUIPlugin);
    univer.registerPlugin(UniverSheetsNumfmtPlugin);
    univer.registerPlugin(UniverSheetsNumfmtUIPlugin);

    // Register the lazy-loader's holder so the eager/idle loaders can reach this
    // univer. CasualSheets uses its own (bundled) copy of the loader — the
    // exported `@casualoffice/sheets/univer` is @internal and only the host app's
    // legacy UniverSheet consumes it, so there's no cross-instance state to share.
    if (lazyPlugins) setUniverForLazyLoad(univer);

    let cancelled = false;
    let changeTimer: ReturnType<typeof setTimeout> | null = null;
    let changeSub: { dispose: () => void } | undefined;

    void (async () => {
      // Eager-load any feature plugin whose data already lives in initialData
      // (CF rules, tables, hyperlinks, …) BEFORE createUnit — Univer's resource
      // manager silently drops keys for plugins that aren't registered when it
      // reads the snapshot. Skipped entirely when lazyPlugins is false.
      if (lazyPlugins) {
        await eagerLoadForSnapshot(univer, initialData);
        if (cancelled) return;
      }

      univer.createUnit(UniverInstanceType.UNIVER_SHEET, initialData);

      const api = createCasualSheetsAPI(FUniver.newAPI(univer));
      apiRef.current = api;
      // Hand the live API to the built-in chrome (FormulaBar subscribes to it).
      // Only when chrome is shown, so bare-grid consumers never re-render.
      if (!cancelled && chrome !== 'none') setChromeApi(api);
      // Apply the initial appearance now that the editor exists (the reactive
      // effect below also runs on mount, but apiRef may not be set yet when it
      // first fires — this guarantees dark mode from the first paint).
      applyAppearance(api, container, appearance);
      onReady?.(api);

      // Debounced snapshot stream → onChange. Subscribed AFTER createUnit so the
      // initial unit-creation mutations don't fire a spurious first emit. Uses the
      // mutation hook (CLAUDE.md hard rule), never UI events.
      if (hasOnChange) {
        const injector = (api.univer as unknown as { _injector?: { get(t: unknown): unknown } })
          ._injector;
        const cmdSvc = injector?.get(ICommandService) as
          | {
              onMutationExecutedForCollab: (
                l: (info: ICommandInfo, options?: IExecutionOptions) => void,
              ) => { dispose: () => void };
            }
          | undefined;
        changeSub = cmdSvc?.onMutationExecutedForCollab(() => {
          if (changeTimer) clearTimeout(changeTimer);
          changeTimer = setTimeout(() => {
            const snap = api.getSnapshot();
            if (snap) onChangeRef.current?.(snap);
          }, onChangeDebounceMs);
        });
        // If we unmounted during the eager-load await, cleanup already ran with
        // changeSub still undefined — dispose this late subscription.
        if (cancelled) changeSub?.dispose();
      }

      // Idle-load the remaining feature plugins so Insert / Data / Format actions
      // are ready when the user reaches them.
      if (lazyPlugins) idleLoadAll(univer);
    })();

    return () => {
      cancelled = true;
      if (changeTimer) clearTimeout(changeTimer);
      changeSub?.dispose();
      apiRef.current = null;
      setChromeApi(null);
      if (lazyPlugins) setUniverForLazyLoad(null);
      // Defer disposal off the React render phase — Univer owns its
      // own React root, and a synchronous unmount mid-render warns
      // and leaves the canvas detached.
      const toDispose = univer;
      queueMicrotask(() => toDispose.dispose());
    };
    // initialData is intentionally NOT in the dep array — the wrapper
    // mounts the snapshot once. Hosts that need to swap workbooks
    // change the React `key` to force a remount.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Reactive appearance. Runs after the boot effect (apiRef populated on first
  // mount) and re-runs whenever `appearance` flips, re-theming the live editor.
  useEffect(() => {
    const api = apiRef.current;
    const container = hostRef.current;
    if (!api || !container) return;
    applyAppearance(api, container, appearance);
  }, [appearance]);

  // chrome="none" (default) keeps the exact bare-grid shape existing consumers
  // rely on (embed-runtime, hosts that bring their own shell). Any other level
  // wraps the grid in a flex column with the built-in chrome above it; the grid
  // container (hostRef, where Univer mounts) fills the remaining space.
  if (chrome === 'none') {
    return (
      <div
        ref={hostRef}
        style={{ ...DEFAULT_STYLE, ...style }}
        className={className}
        data-testid={testId}
      />
    );
  }

  return (
    <div
      className={className}
      data-testid={testId}
      style={{ ...DEFAULT_STYLE, ...style, display: 'flex', flexDirection: 'column' }}
    >
      <Toolbar getApi={() => apiRef.current} />
      <FormulaBar api={chromeApi} />
      <div ref={hostRef} style={{ flex: '1 1 auto', minHeight: 0, position: 'relative' }} />
      <StatusBar api={chromeApi} />
    </div>
  );
}

/**
 * Apply light/dark to a live editor. `ThemeService.setDarkMode` is the source of
 * truth — it flips the canvas colours, the internals that subscribe to
 * `darkMode$` (notifications, message containers), AND Univer's Workbench toggles
 * the `univer-dark` class on the document root for its compiled dark CSS. We also
 * mirror the class onto the editor container as race-insurance (the Workbench
 * effect can land a frame after ours). Mirrors the app's ThemeBridge.
 */
function applyAppearance(
  api: CasualSheetsAPI,
  container: HTMLElement,
  appearance: 'light' | 'dark',
): void {
  const dark = appearance === 'dark';
  container.classList.toggle('univer-dark', dark);
  try {
    const injector = (api.univer as unknown as { _injector?: { get(t: unknown): unknown } })
      ._injector;
    const themeService = injector?.get(ThemeService) as
      | { setDarkMode(b: boolean): void; darkMode: boolean }
      | undefined;
    if (themeService && themeService.darkMode !== dark) themeService.setDarkMode(dark);
  } catch {
    /* ThemeService unavailable — the class toggle still themes visible chrome */
  }
}
