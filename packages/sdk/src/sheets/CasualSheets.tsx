/**
 * CasualSheets — minimal React wrapper around Univer Sheets.
 *
 * Boots Univer with the eager plugin set (render + formula engine +
 * UI + docs + sheets + sheets-ui + sheets-formula + numfmt), mounts a
 * single workbook unit from `initialData`, and hands the host the
 * `CasualSheetsAPI` imperative ref via `onReady` (raw FUniver facade
 * available at `api.univer`).
 *
 * Intentionally NOT included (host can layer on top via FUniver):
 *   - Lazy plugin loading (conditional formatting, drawings, sort,
 *     filter, hyperlinks, tables, comments, find/replace, …). Host
 *     calls `univer.registerPlugin(...)` after `onReady`.
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

import { useEffect, useRef, type CSSProperties } from 'react';
import {
  ICommandService,
  LocaleType,
  LogLevel,
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
  /** Override the theme. Defaults to Univer's `defaultTheme`. */
  theme?: typeof defaultTheme;
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
  locale = LocaleType.EN_US,
  locales,
  logLevel = LogLevel.WARN,
  ui,
  theme = defaultTheme,
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

    univer.createUnit(UniverInstanceType.UNIVER_SHEET, initialData);

    const api = createCasualSheetsAPI(FUniver.newAPI(univer));
    onReady?.(api);

    // Debounced snapshot stream → onChange. Subscribed AFTER createUnit so the
    // initial unit-creation mutations don't fire a spurious first emit. Uses the
    // mutation hook (CLAUDE.md hard rule), never UI events.
    let changeTimer: ReturnType<typeof setTimeout> | null = null;
    let changeSub: { dispose: () => void } | undefined;
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
    }

    return () => {
      if (changeTimer) clearTimeout(changeTimer);
      changeSub?.dispose();
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

  return (
    <div
      ref={hostRef}
      style={{ ...DEFAULT_STYLE, ...style }}
      className={className}
      data-testid={testId}
    />
  );
}
