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
 * Formula compute runs on the main thread by default; pass `formula={{ worker }}`
 * to move it off-thread (the SDK then wires `UniverRPCMainThreadPlugin` to the
 * host's worker — see the `formula` prop).
 *
 * Intentionally NOT included (host can layer on top via FUniver):
 *   - Snapshot swap (this component mounts a snapshot once; change
 *     the React `key` to remount with a fresh snapshot).
 *   - Paste-merge hooks, dev helpers, zoom-shortcut overrides,
 *     facade extensions — all app concerns.
 *
 * Styles: host must import `@casualoffice/sheets/styles.css`
 * (or the per-plugin CSS) once at app boot. Tree-shaking strips the
 * styles from this entry if the host doesn't reach the styles export.
 */

import {
  lazy,
  Suspense,
  useEffect,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent,
} from 'react';
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
import { UniverSheetsFormulaPlugin, CalculationMode } from '@univerjs/sheets-formula';
import { UniverSheetsFormulaUIPlugin } from '@univerjs/sheets-formula-ui';
import { UniverSheetsNumfmtPlugin } from '@univerjs/sheets-numfmt';
import { UniverSheetsNumfmtUIPlugin } from '@univerjs/sheets-numfmt-ui';
// Type-only — erased at build, so `@univerjs/rpc` stays a runtime-optional peer
// (loaded via dynamic import only when a formula worker is passed).
import type { UniverRPCMainThreadPlugin as RpcMainThreadPluginType } from '@univerjs/rpc';

import { createCasualSheetsAPI, type CasualSheetsAPI } from './api';
import {
  eagerLoadForSnapshot,
  ensurePlugin,
  idleLoadAll,
  setUniverForLazyLoad,
} from '../univer/lazy-plugins';
import type { ChromeExtensions } from '../chrome/extensions';
import type { DialogKind } from '../chrome/dialog-context';
// Chrome is lazy-loaded from the `@casualoffice/sheets/chrome` subpath (NOT a
// relative import — that would inline under this build's splitting:false). The
// subpath is externalised in tsup, so the consumer's bundler code-splits it and
// `chrome="none"` hosts (the default + the apps/web reference host) never load
// the chrome chunk.
const ChromeTop = lazy(() =>
  import('@casualoffice/sheets/chrome').then((m) => ({ default: m.ChromeTop })),
);
const ChromeBottom = lazy(() =>
  import('@casualoffice/sheets/chrome').then((m) => ({ default: m.ChromeBottom })),
);

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
  /** Explicit save — fired when the user presses Ctrl/Cmd+S inside the editor
   *  (the browser's save dialog is suppressed). The host persists the snapshot.
   *  Part of the "host owns storage" contract: the SDK never writes a store. */
  onSave?: (snapshot: IWorkbookData) => void;
  /** Fired once when the editor unmounts, with the final snapshot — the host's
   *  last chance to persist before the workbook is disposed. */
  onExit?: (snapshot: IWorkbookData) => void;
  /** Lazy-load the feature plugins (conditional formatting, data
   *  validation, hyperlinks, notes, tables, comments, drawings, sort,
   *  filter, find/replace). Default `true`: plugins whose data is in
   *  `initialData` load eagerly before mount (so nothing is dropped on
   *  open), the rest idle-load after first paint. Set `false` for the
   *  minimal editor (render + formula + numfmt only) — the embed-iframe
   *  build does this to stay a single self-contained bundle. */
  lazyPlugins?: boolean;
  /** Escape hatch fired after the SDK registers its built-in plugins but BEFORE
   *  the workbook unit is created — the host can `univer.registerPlugin(...)`
   *  additional plugins here (e.g. an off-main formula worker via
   *  `UniverRPCMainThreadPlugin`, crosshair-highlight, zen-editor). Anything
   *  registered after `createUnit` would miss the unit's plugin-init pass, so
   *  register-time extras must go through this hook. Power hosts (the reference
   *  app) use it to share the SDK editor core while keeping their extra plugins;
   *  most integrators never need it. NOT covered by semver — it hands you the
   *  raw `Univer` instance. */
  onBeforeCreateUnit?: (univer: Univer) => void;
  /** Off-main formula compute. By default the formula engine runs on the MAIN
   *  thread (fine for typical sheets, zero host setup). Provide a Web Worker (or
   *  its URL) to move compute off-thread so paste / sort / fill on large
   *  workbooks don't freeze the UI: the SDK then registers the formula plugins
   *  with `notExecuteFormula` and wires `UniverRPCMainThreadPlugin` to your
   *  worker. The host owns the worker (the SDK never bundles one — that's brittle
   *  across bundlers) and must have `@univerjs/rpc` installed. The worker script
   *  is the standard Univer formula worker (see the reference app's
   *  `apps/web/src/univer/formula-worker.ts`). */
  formula?: {
    /** A constructed `Worker`, or a URL/string the RPC plugin loads. */
    worker?: Worker | string;
  };
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
  /** Office chrome rendered around the grid:
   *  - `'none'` (default): bare grid — the host supplies its own chrome.
   *  - `'minimal'` / `'full'`: the built-in Office shell — a menu bar
   *    (Edit/Insert/Format/Data/View), a formatting toolbar (font family/size,
   *    bold/italic/underline/strike, text & fill colour, borders, h/v align,
   *    wrap, merge, number formats, clear format, AutoSum), a formula bar with a
   *    name box + function autocomplete, a worksheet tab strip (switch/add/
   *    rename/delete), and a status bar (Average/Count/Numerical Count/Min/Max/
   *    Sum + zoom). All driven through the facade, themed via `--cs-chrome-*`
   *    (light/dark). `'minimal'` and `'full'` currently render the same shell;
   *    `'full'` is where richer panels (find/replace, charts, …) will land. */
  chrome?: 'none' | 'minimal' | 'full';
  /** Enable/disable chrome features. Each key maps a toolbar group / menu item /
   *  capability to a boolean; `false` hides the control AND blocks its command.
   *  Omitted keys default to enabled. Only applies when `chrome` is shown. */
  features?: Record<string, boolean>;
  /** Legacy host hook for dialog-backed chrome controls. The SDK now ships
   *  BUILT-IN dialogs (Format Cells, Find & Replace, …) that open by default, so
   *  this is no longer required. It still works for back-compat: kinds the SDK
   *  has no built-in for (Insert Chart, PivotTable, …) fall through to it, and a
   *  host can force specific kinds to it via `hostOwnedDialogs`. Prefer
   *  `extensions.dialogs` to supply a React override component. */
  onDialogRequest?: (kind: string, context?: unknown) => void;
  /** Kinds the host wants to handle via `onDialogRequest` even though the SDK has
   *  a built-in (e.g. keep the SDK chrome but your own Format Cells). */
  hostOwnedDialogs?: DialogKind[];
  /** Host chrome extensions — the extensibility surface for `chrome="full"`.
   *  Append custom toolbar items / menu items / side panels, and register or
   *  OVERRIDE dialogs by kind. Built-ins are the defaults; host entries
   *  append/override. See `ChromeExtensions` for the exact shape. */
  extensions?: ChromeExtensions;
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
  onSave,
  onExit,
  lazyPlugins = true,
  onBeforeCreateUnit,
  formula,
  locale = LocaleType.EN_US,
  locales,
  logLevel = LogLevel.WARN,
  ui,
  theme = defaultTheme,
  appearance = 'light',
  chrome = 'none',
  features,
  onDialogRequest,
  hostOwnedDialogs,
  extensions,
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
  // Latest save/exit callbacks, called via refs so they fire without re-running
  // the boot effect. onExit is read in cleanup; onSave on the Ctrl/Cmd+S handler.
  const onSaveRef = useRef(onSave);
  onSaveRef.current = onSave;
  const onExitRef = useRef(onExit);
  onExitRef.current = onExit;
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

    // `formula.worker` → off-main compute. Default = main thread (fine for
    // typical sheets, zero host setup).
    const offMain = !!formula?.worker;

    let cancelled = false;
    let changeTimer: ReturnType<typeof setTimeout> | null = null;
    let changeSub: { dispose: () => void } | undefined;

    void (async () => {
      // Plugin registration runs here (not synchronously) so the OPTIONAL RPC
      // transport can be `await import`ed FIRST and registered in its correct
      // slot — right after the formula engine, before sheets. Registering it out
      // of order (or after createUnit) leaves the formula engine's worker channel
      // unwired → cells stay 0. Dynamic import keeps `@univerjs/rpc` a true
      // optional peer (only loaded when a worker is passed).
      let RPCMainThreadPlugin: typeof RpcMainThreadPluginType | null = null;
      if (offMain && formula?.worker) {
        RPCMainThreadPlugin = (await import('@univerjs/rpc')).UniverRPCMainThreadPlugin;
        if (cancelled) return;
      }

      univer.registerPlugin(UniverRenderEnginePlugin);
      univer.registerPlugin(
        UniverFormulaEnginePlugin,
        offMain ? { notExecuteFormula: true } : undefined,
      );
      if (RPCMainThreadPlugin && formula?.worker) {
        univer.registerPlugin(RPCMainThreadPlugin, { workerURL: formula.worker });
      }
      univer.registerPlugin(UniverUIPlugin, uiOpts);
      univer.registerPlugin(UniverDocsPlugin);
      univer.registerPlugin(UniverDocsUIPlugin);
      univer.registerPlugin(UniverSheetsPlugin, offMain ? { notExecuteFormula: true } : undefined);
      univer.registerPlugin(UniverSheetsUIPlugin);
      univer.registerPlugin(
        UniverSheetsFormulaPlugin,
        offMain
          ? { notExecuteFormula: true, initialFormulaComputing: CalculationMode.NO_CALCULATION }
          : undefined,
      );
      univer.registerPlugin(UniverSheetsFormulaUIPlugin);
      univer.registerPlugin(UniverSheetsNumfmtPlugin);
      univer.registerPlugin(UniverSheetsNumfmtUIPlugin);

      // Lazy-loader holder (the loader is @internal so a relative import shares
      // no cross-instance state) + host plugin escape hatch — both before
      // createUnit.
      if (lazyPlugins) setUniverForLazyLoad(univer);
      onBeforeCreateUnit?.(univer);

      // Eager-load any feature plugin whose data already lives in initialData
      // (CF rules, tables, hyperlinks, …) BEFORE createUnit — Univer's resource
      // manager silently drops keys for plugins that aren't registered when it
      // reads the snapshot. Skipped entirely when lazyPlugins is false.
      if (lazyPlugins) {
        await eagerLoadForSnapshot(univer, initialData);
        if (cancelled) return;
        // Drawing/image is the one feature whose trigger (Insert ▸ Image) opens
        // a FILE PICKER — which needs the user's click gesture. If the plugin
        // lazy-loads on click, the await loses the gesture and the picker never
        // opens ("can't insert image"); if it idle-loads, a quick click before
        // it's ready silently no-ops. So load it eagerly here (tracked by
        // ensurePlugin, so idleLoadAll won't double-register) — image works on
        // the first click, in-gesture. Other features open panels (no gesture).
        await ensurePlugin(univer, 'drawing');
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
      // Last-chance persist: emit the final snapshot before the workbook is
      // disposed (disposal is deferred via microtask below, so it's still alive).
      if (onExitRef.current) {
        const snap = apiRef.current?.getSnapshot();
        if (snap) onExitRef.current(snap);
      }
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

  // Ctrl/Cmd+S anywhere in the editor → onSave (suppress the browser dialog).
  // Capture phase so we beat Univer's own key handling on the canvas.
  const onKeyDownCapture = (e: ReactKeyboardEvent<HTMLDivElement>) => {
    if ((e.metaKey || e.ctrlKey) && (e.key === 's' || e.key === 'S')) {
      e.preventDefault();
      const snap = apiRef.current?.getSnapshot();
      if (snap) onSaveRef.current?.(snap);
    }
  };

  // chrome="none" (default) keeps the exact bare-grid shape existing consumers
  // rely on (embed-runtime, hosts that bring their own shell). Any other level
  // wraps the grid in a flex column with the built-in chrome above it; the grid
  // container (hostRef, where Univer mounts) fills the remaining space.
  if (chrome === 'none') {
    return (
      <div
        ref={hostRef}
        onKeyDownCapture={onKeyDownCapture}
        style={{ ...DEFAULT_STYLE, ...style }}
        className={className}
        data-testid={testId}
      />
    );
  }

  // The built-in chrome components read their colours from `--cs-chrome-*` CSS
  // vars. Phase 4: each var now resolves to a `@schnsrw/design-system` token
  // (loaded by the host via `tokens.css`), with the prior hardcoded value as a
  // FALLBACK so the chrome still renders standalone for hosts that don't ship the
  // design system. `data-theme` on the wrapper (below) swaps the DS tokens
  // light/dark; the fallbacks keep `appearance` working without the DS too.
  const dark = appearance === 'dark';
  const chromeVars = {
    '--cs-chrome-bg': `var(--color-surface-strip, ${dark ? '#2a2e35' : '#eef1f5'})`,
    '--cs-chrome-fg': `var(--color-text, ${dark ? '#e6e6e6' : '#201f1e'})`,
    '--cs-chrome-muted': `var(--color-text-secondary, ${dark ? '#b0b3ba' : '#605e5c'})`,
    '--cs-chrome-border': `var(--color-divider, ${dark ? '#24272d' : '#edeff3'})`,
    '--cs-chrome-input-bg': `var(--color-surface, ${dark ? '#1b1e23' : '#ffffff'})`,
    '--cs-chrome-hover': `var(--color-hover, ${dark ? 'rgba(255,255,255,0.06)' : 'rgba(15,23,42,0.045)'})`,
    '--cs-chrome-active': `var(--color-selected, ${dark ? 'rgba(34,211,238,0.20)' : 'rgba(14,116,144,0.11)'})`,
    '--cs-chrome-active-fg': `var(--color-accent, ${dark ? '#22d3ee' : '#0e7490'})`,
  } as CSSProperties;

  return (
    <div
      className={className}
      data-testid={testId}
      data-theme={dark ? 'dark' : 'light'}
      onKeyDownCapture={onKeyDownCapture}
      style={{
        ...DEFAULT_STYLE,
        ...chromeVars,
        ...style,
        display: 'flex',
        flexDirection: 'column',
      }}
    >
      {/* Bars appear once their lazy chunk loads (a tick after first paint); the
          grid host is OUTSIDE Suspense so Univer mounts immediately. */}
      <Suspense fallback={null}>
        <ChromeTop
          api={chromeApi}
          features={features}
          onDialogRequest={onDialogRequest}
          hostOwnedDialogs={hostOwnedDialogs}
          extensions={extensions}
        />
      </Suspense>
      <div ref={hostRef} style={{ flex: '1 1 auto', minHeight: 0, position: 'relative' }} />
      <Suspense fallback={null}>
        <ChromeBottom api={chromeApi} />
      </Suspense>
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
