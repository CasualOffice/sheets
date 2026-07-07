/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

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
  IMentionIOService,
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
import { CasualMentionIOService } from './mention-io';
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

import {
  createCasualSheetsAPI,
  type CasualSheetsAPI,
  type CasualSheetsAPIInternal,
  type DocumentMode,
  type RangeRef,
} from './api';
// Type-only — erased at build, so the collab entry (Yjs + Hocuspocus) stays out
// of the `sheets` bundle. The runtime `attachCollab` is pulled in lazily via a
// dynamic `import('@casualoffice/sheets/collab')` only when the `collab` prop is
// set (see the collab effect), so bare-grid hosts never load it.
import type { AttachCollabOptions, CollabHandle } from '../collab/attachCollab';
import {
  eagerLoadForSnapshot,
  ensurePlugin,
  idleLoadAll,
  setUniverForLazyLoad,
} from '../univer/lazy-plugins';
import type { ChromeExtensions } from '../chrome/extensions';
import type { DialogKind } from '../chrome/dialog-context';
import { AiPanelSurface, type SheetsAiConfig } from '../ai/AiPanelSurface';
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
  /** The active selection changed (canvas-driven), or `null` when there is none.
   *  The prop half of the canonical `selectionChange` event (doc 38 §3); the
   *  same event is available via `api.on('selectionChange', …)`. Wired to
   *  Univer's `SelectionChanged` / `SelectionMoveEnd` facade events. */
  onSelectionChange?: (selection: RangeRef | null) => void;
  /** A boot/runtime error surfaced by the editor — the prop half of the
   *  canonical `error` event (doc 38 §3). Also available via
   *  `api.on('error', …)`. */
  onError?: (error: Error) => void;
  /** The unsaved-changes flag flipped: `true` after the first edit since the
   *  last load/save, `false` on save / `setContent` / `import`. The prop half of
   *  the canonical `dirtyChange` event (doc 38 §3); also `api.on('dirtyChange',
   *  …)`. Lets a host drive a "•/unsaved" title dot without diffing snapshots. */
  onDirtyChange?: (dirty: boolean) => void;
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
  /** Document interaction mode (SuperDoc-aligned vocabulary, shared with the
   *  docs SDK):
   *  - `'editing'` (default): fully editable.
   *  - `'viewing'`: read-only — applies the command-veto + permission path
   *    (`applyReadOnly`) to the mounted workbook.
   *  Reactive: flipping it re-applies via `api.setDocumentMode`. Wins over the
   *  deprecated `readOnly` prop when both are set. */
  documentMode?: DocumentMode;
  /** Real-time co-editing, declaratively. Pass `{ server, room, … }` to join a
   *  room and the SDK wires Yjs/Hocuspocus itself once the editor is ready
   *  (mirrors how the docs `CasualEditor` opts in via `backendUrl`); omit it for
   *  a single-user editor. Options match `attachCollab` (`server`, `room`,
   *  `password`, `token`, `role`, `share`, `onStatus`, `onSnapshot`).
   *
   *  Lifecycle: attaches after `onReady`, detaches on unmount, and re-attaches
   *  when `server` / `room` / `password` / `token` / `role` change. The
   *  imperative `attachCollab(api, opts)` export stays available for advanced
   *  hosts that drive the room lifecycle themselves (presence UI, preflight,
   *  reconnect banners) — don't combine both on one editor. */
  collab?: AttachCollabOptions;
  /** AI assistant surface, declaratively. Pass `{ enabled: true, transport,
   *  render }` to mount a supported AI task-pane beside the grid — mirrors how
   *  the docs SDK exposes its DocOps panel behind a transport prop. The SDK
   *  owns the prop contract, the `SheetsAiTransport` type, and the layout slot;
   *  the panel body is supplied via `ai.render` (the reference app passes its
   *  `<AiPanel>`) and drives its tool loop against `ai.transport`. Build a
   *  transport with `createSheetsAiTransport()` (desktop-native → collab
   *  single-round → browser-direct). Omit `ai` for an editor with no AI.
   *
   *  Reactive: flipping `enabled` mounts/unmounts the pane once the editor is
   *  ready. See `SheetsAiConfig`. */
  ai?: SheetsAiConfig;
  /** @deprecated Use `documentMode` instead. `true` maps to
   *  `documentMode="viewing"`. Ignored when `documentMode` is set. */
  readOnly?: boolean;
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
  onSelectionChange,
  onError,
  onDirtyChange,
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
  documentMode,
  collab,
  ai,
  readOnly,
  style,
  className,
  testId = 'casual-sheets',
}: CasualSheetsProps) {
  // `documentMode` wins; the deprecated `readOnly` boolean only applies when
  // `documentMode` is unset. Absent both → editable.
  const effectiveMode: DocumentMode = documentMode ?? (readOnly ? 'viewing' : 'editing');
  const hostRef = useRef<HTMLDivElement>(null);
  // Keep the latest onChange callable without re-subscribing (the effect
  // mounts once). The mutation subscription is always wired (it also backs the
  // `dirtyChange` / `change` emitter); only the snapshot serialization inside it
  // is gated on an actual consumer.
  const onChangeRef = useRef(onChange);
  onChangeRef.current = onChange;
  // Latest save/exit callbacks, called via refs so they fire without re-running
  // the boot effect. onExit is read in cleanup; onSave on the Ctrl/Cmd+S handler.
  const onSaveRef = useRef(onSave);
  onSaveRef.current = onSave;
  const onExitRef = useRef(onExit);
  onExitRef.current = onExit;
  // Latest handlers for the canonical event props (doc 38 §3), read via refs so
  // they fire without re-running the boot effect. `hasX` (fixed at mount) gates
  // whether we bridge each prop to the unified emitter — a host that passes none
  // of them (and never calls `api.on`) pays for none of the extra wiring.
  const onSelectionChangeRef = useRef(onSelectionChange);
  onSelectionChangeRef.current = onSelectionChange;
  const onErrorRef = useRef(onError);
  onErrorRef.current = onError;
  const onDirtyChangeRef = useRef(onDirtyChange);
  onDirtyChangeRef.current = onDirtyChange;
  const hasSelectionChange = useRef(!!onSelectionChange).current;
  const hasError = useRef(!!onError).current;
  const hasDirtyChange = useRef(!!onDirtyChange).current;
  // The live FUniver facade, captured at mount so the reactive appearance
  // effect can reach Univer's ThemeService without re-running boot.
  const apiRef = useRef<CasualSheetsAPI | null>(null);
  // The live API as state, so the built-in chrome (FormulaBar) re-renders and
  // subscribes once the editor is ready. Only set when chrome is shown — the
  // bare-grid path never triggers this re-render. A single post-mount setState
  // doesn't disturb the grid (Univer owns its canvas outside React).
  const [chromeApi, setChromeApi] = useState<CasualSheetsAPI | null>(null);
  // Declarative collab. `hasCollab` is fixed at mount (the boot effect runs
  // once): only then do we surface the ready api as state to drive the attach
  // effect, so single-user consumers never take the extra re-render. The latest
  // options live on a ref so callback/`share` changes don't force a re-attach —
  // the attach effect only re-runs on connection-identity changes.
  const hasCollab = useRef(!!collab).current;
  const collabRef = useRef(collab);
  collabRef.current = collab;
  const [collabApi, setCollabApi] = useState<CasualSheetsAPI | null>(null);
  // AI surface. Like chrome/collab, only surface the ready api as state when
  // the `ai` prop was present at mount — so editors with no AI never take the
  // extra re-render. The panel itself only mounts when `ai.enabled` is set.
  const hasAi = useRef(!!ai).current;
  const [aiApi, setAiApi] = useState<CasualSheetsAPI | null>(null);

  useEffect(() => {
    const container = hostRef.current;
    if (!container) return;

    const univer = new Univer({
      theme,
      locale,
      locales,
      logLevel,
      // Replace the default mention IO (hardwired to the current user) with our
      // host-pluggable source so comment @-mentions can list real collaborators.
      // No-op until a provider is installed via `setMentionProvider`.
      override: [[IMentionIOService, { useClass: CasualMentionIOService }]],
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
      const apiInternal = api as CasualSheetsAPIInternal;
      apiRef.current = api;
      // Bridge the declarative event props (doc 38 §3) to the unified emitter, so
      // a prop and `api.on(name, …)` both receive the event. Only wire the bridge
      // when the prop was present at mount — direct `api.on(...)` still works
      // without it (the emitter fires to its own subscribers regardless). The
      // selection bridge is what makes the factory's `SelectionChanged` listener
      // do real work, so bare hosts that pass no `onSelectionChange` (and never
      // call `api.on`) never pay for `getSelection` on every selection move.
      if (hasSelectionChange) {
        apiInternal.on('selectionChange', (sel) => onSelectionChangeRef.current?.(sel));
      }
      if (hasError) {
        apiInternal.on('error', (err) => onErrorRef.current?.(err));
      }
      if (hasDirtyChange) {
        apiInternal.on('dirtyChange', (d) => onDirtyChangeRef.current?.(d));
      }
      // Hand the live API to the built-in chrome (FormulaBar subscribes to it).
      // Only when chrome is shown, so bare-grid consumers never re-render.
      if (!cancelled && chrome !== 'none') setChromeApi(api);
      // Hand the ready api to the declarative-collab effect (only when the
      // `collab` prop was present at mount — otherwise no extra re-render).
      if (!cancelled && hasCollab) setCollabApi(api);
      // Hand the ready api to the AI surface (only when `ai` was present at
      // mount — otherwise no extra re-render).
      if (!cancelled && hasAi) setAiApi(api);
      // Apply the initial appearance now that the editor exists (the reactive
      // effect below also runs on mount, but apiRef may not be set yet when it
      // first fires — this guarantees dark mode from the first paint).
      applyAppearance(api, container, appearance);
      // Apply the initial document mode. Deferred a frame for the read-only case:
      // sheets-ui sets WorkbookEditablePermission → true during unit setup AFTER
      // onReady, which would clobber a synchronous flip (the embed preview path
      // waits a rAF for the same reason). The command-veto layer is unaffected,
      // but the rAF keeps the permission layer in sync too.
      if (effectiveMode === 'viewing') {
        requestAnimationFrame(() => {
          if (!cancelled) api.setDocumentMode('viewing');
        });
      }
      onReady?.(api);
      // Fire the canonical `ready` event for `api.on('ready', …)` subscribers.
      // `ready` is sticky in the emitter, so a late subscription still fires.
      if (!cancelled) apiInternal.emit('ready', api);

      // Mutation stream → dirty flag + debounced `change` (prop + emitter).
      // Subscribed AFTER createUnit so the initial unit-creation mutations don't
      // fire a spurious first emit. Uses the mutation hook (CLAUDE.md hard rule),
      // never UI events. Always subscribed now (not just when `onChange` was
      // passed) because `dirtyChange` and `api.on('change', …)` also depend on
      // it; the raw callback is cheap (a dirty-flag flip + a debounce timer), and
      // the expensive snapshot serialization is gated on an actual consumer.
      {
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
          // First edit since load/save → dirty (emits dirtyChange on transition).
          apiInternal.markDirty(true);
          if (changeTimer) clearTimeout(changeTimer);
          changeTimer = setTimeout(() => {
            // Only serialize when something actually consumes the snapshot.
            if (!onChangeRef.current && apiInternal.listenerCount('change') === 0) return;
            const snap = api.getContent();
            if (snap) {
              onChangeRef.current?.(snap);
              apiInternal.emit('change', snap);
            }
          }, onChangeDebounceMs);
        });
        // If we unmounted during the eager-load await, cleanup already ran with
        // changeSub still undefined — dispose this late subscription.
        if (cancelled) changeSub?.dispose();
      }

      // Idle-load the remaining feature plugins so Insert / Data / Format actions
      // are ready when the user reaches them.
      if (lazyPlugins) idleLoadAll(univer);
    })().catch((err: unknown) => {
      // Boot failure (plugin load, unit creation, …) → the canonical `error`
      // event. Prefer the emitter once the API exists (reaches both the `onError`
      // prop bridge and `api.on('error', …)`); before then fall back to the prop.
      if (cancelled) return;
      const e = err instanceof Error ? err : new Error(String(err));
      const cur = apiRef.current as CasualSheetsAPIInternal | null;
      if (cur) cur.emit('error', e);
      else onErrorRef.current?.(e);
    });

    return () => {
      cancelled = true;
      if (changeTimer) clearTimeout(changeTimer);
      changeSub?.dispose();
      // Last-chance persist: emit the final snapshot before the workbook is
      // disposed (disposal is deferred via microtask below, so it's still alive).
      if (onExitRef.current) {
        const snap = apiRef.current?.getContent();
        if (snap) onExitRef.current(snap);
      }
      apiRef.current = null;
      setChromeApi(null);
      setCollabApi(null);
      setAiApi(null);
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

  // Reactive document mode. On first mount apiRef may not be populated yet (boot
  // is async) — that run bails and the boot effect applies the initial mode; this
  // effect then handles every subsequent `documentMode` / `readOnly` flip.
  useEffect(() => {
    apiRef.current?.setDocumentMode(effectiveMode);
  }, [effectiveMode]);

  // Declarative collab: attach once the api is ready, detach on unmount, and
  // re-attach when the connection identity changes. `attachCollab` (and its
  // Yjs/Hocuspocus transport) is loaded lazily via the externalised
  // `@casualoffice/sheets/collab` subpath — same code-split as `chrome`/`xlsx`
  // — so single-user hosts never pull it into their bundle. Options are read
  // from `collabRef` at attach time so callback/`share` changes don't re-attach.
  useEffect(() => {
    if (!collabApi || !collabRef.current) return;
    let handle: CollabHandle | null = null;
    let disposed = false;
    void import('@casualoffice/sheets/collab').then(({ attachCollab }) => {
      const opts = collabRef.current;
      if (disposed || !opts) return;
      handle = attachCollab(collabApi, opts);
    });
    return () => {
      disposed = true;
      handle?.detach();
      handle = null;
    };
    // Re-attach only on connection-identity changes (server/room/auth/role); the
    // options object identity and callbacks are read live from `collabRef`.
  }, [collabApi, collab?.server, collab?.room, collab?.password, collab?.token, collab?.role]);

  // Ctrl/Cmd+S anywhere in the editor → onSave (suppress the browser dialog).
  // Capture phase so we beat Univer's own key handling on the canvas.
  const onKeyDownCapture = (e: ReactKeyboardEvent<HTMLDivElement>) => {
    if ((e.metaKey || e.ctrlKey) && (e.key === 's' || e.key === 'S')) {
      e.preventDefault();
      const api = apiRef.current;
      const snap = api?.getContent();
      if (snap) {
        onSaveRef.current?.(snap);
        // Canonical `save` event + clear the dirty flag (an explicit save means
        // the buffer is clean until the next edit).
        const apiInternal = api as CasualSheetsAPIInternal;
        apiInternal.emit('save', snap);
        apiInternal.markDirty(false);
      }
    }
  };

  // chrome="none" (default) keeps the exact bare-grid shape existing consumers
  // rely on (embed-runtime, hosts that bring their own shell). Any other level
  // wraps the grid in a flex column with the built-in chrome above it; the grid
  // container (hostRef, where Univer mounts) fills the remaining space.
  if (chrome === 'none') {
    // No `ai` prop → the exact bare-grid shape existing consumers rely on.
    if (!hasAi) {
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
    // With `ai`, the grid + AI pane sit in a stable flex row (shape fixed at
    // mount by `hasAi`, so toggling `ai.enabled` never remounts the grid host —
    // it only adds/removes the aside sibling). The grid host still fills the
    // remaining space, so Univer's canvas sizing is unchanged.
    return (
      <div
        className={className}
        data-testid={testId}
        onKeyDownCapture={onKeyDownCapture}
        style={{ ...DEFAULT_STYLE, ...style, display: 'flex', flexDirection: 'row' }}
      >
        <div
          ref={hostRef}
          style={{ flex: '1 1 auto', minWidth: 0, minHeight: 0, position: 'relative' }}
        />
        <AiPanelSurface config={ai} api={aiApi} />
      </div>
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
      {hasAi ? (
        // Grid + AI pane share a flex row so the built-in chrome (top/bottom
        // bars) still spans the full width. Shape fixed at mount by `hasAi`.
        <div style={{ flex: '1 1 auto', minHeight: 0, display: 'flex', flexDirection: 'row' }}>
          <div
            ref={hostRef}
            style={{ flex: '1 1 auto', minWidth: 0, minHeight: 0, position: 'relative' }}
          />
          <AiPanelSurface config={ai} api={aiApi} />
        </div>
      ) : (
        <div ref={hostRef} style={{ flex: '1 1 auto', minHeight: 0, position: 'relative' }} />
      )}
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
