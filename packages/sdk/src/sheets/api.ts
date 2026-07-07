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
 * CasualSheetsAPI — the imperative ref handed to a host via `<CasualSheets onReady>`.
 *
 * This is the SDK's stable integration surface (Excalidraw's model: props +
 * imperative ref). Hosts drive the editor through these methods rather than
 * reaching into Univer directly; `api.univer` is the documented escape hatch and
 * is explicitly NOT covered by semver — everything else here is.
 *
 * Surface (canonical, doc 38 §4):
 *   getContent / setContent / import / export / getSelection / focus /
 *   on / off / executeCommand / executeCommands / undo / redo / onMutation /
 *   setTheme / setDocumentMode / getDocumentMode / univer
 *   (+ deprecated aliases getSnapshot / loadSnapshot; format-specific
 *   importXlsx / exportXlsx retained)
 *
 * `importXlsx` / `exportXlsx` lazy-load the converters via
 * `import('@casualoffice/sheets/xlsx')` — a BARE subpath, not a relative
 * `import('../xlsx')`. The main tsup config is `splitting:false`, so a relative
 * dynamic import would be inlined and balloon the editor entry from ~24KB to
 * ~200KB of ExcelJS for hosts that never touch a file. The subpath is
 * externalised in tsup.config.ts so it stays a separate chunk the consumer
 * code-splits.
 *
 * `attachCollab` is NOT a method here — it ships as a standalone
 * `attachCollab(api, opts)` on the `@casualoffice/sheets/collab` subpath so the
 * editor stays collab-unaware (and collab-free in the bundle) until opted in.
 */

// Side-effect import: registers the Sheets FUniver mixins
// (getActiveWorkbook / createWorkbook / FWorkbook.save / FRange.getRange / …)
// onto the core FUniver facade — both the runtime methods AND the TypeScript
// type augmentation this file relies on. Without it, FUniver is the bare core
// facade and these methods exist neither at type-check nor at runtime.
import '@univerjs/sheets/facade';
import { ICommandService, IUniverInstanceService, ThemeService } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import type { IRange, IWorkbookData } from '@univerjs/core';
import {
  attachMutationObserver,
  runSteps,
  type CommandRecord,
  type MutationEmitter,
} from './scripting';
import { applyReadOnly } from './read-only';
import { createEmitter } from './emitter';

// Re-export so hosts can type a recorded/scripted step off the main entry.
export type { CommandRecord } from './scripting';

/**
 * Document interaction mode — the SuperDoc-aligned vocabulary shared with the
 * docs SDK so hosts drive both editors the same way:
 *
 *  - `'editing'` — fully editable (the default).
 *  - `'viewing'` — read-only: the same command-veto + permission path as
 *    {@link applyReadOnly}.
 *
 * Sheets has no `'suggesting'` tier (that's a docs-only track-changes concept),
 * so the union is the two states.
 */
export type DocumentMode = 'editing' | 'viewing';

/** The active selection, as a sheet-scoped range. */
export interface RangeRef {
  /** Workbook unit id the selection belongs to. */
  unitId: string;
  /** Worksheet (sub-unit) id the selection belongs to. */
  sheetId: string;
  /** `{ startRow, startColumn, endRow, endColumn }`. */
  range: IRange;
}

/**
 * The canonical cross-editor event map (doc 38 §3), shared in shape with the
 * docs SDK so a host wires both editors the same way. Every event is available
 * two ways: as an `on*` prop on `<CasualSheets>` AND via `api.on(name, handler)` /
 * `api.off(name, handler)` on this handle — same event, same payload.
 *
 * Payloads are the format's own content type where the doc leaves it to the
 * format (`change` / `save` carry an `IWorkbookData` snapshot, not raw bytes).
 */
export interface CasualSheetsEvents {
  /** Fired once the workbook unit is created and the API is live. Sticky: a
   *  handler registered via `api.on('ready', …)` after the editor is already
   *  ready is invoked immediately with the API. */
  ready: (api: CasualSheetsAPI) => void;
  /** Debounced workbook snapshot, emitted after edits settle. */
  change: (snapshot: IWorkbookData) => void;
  /** The active selection changed (canvas-driven). `null` when there is none. */
  selectionChange: (selection: RangeRef | null) => void;
  /** Explicit save (Ctrl/Cmd+S), carrying the current snapshot to persist. */
  save: (snapshot: IWorkbookData) => void;
  /** A boot/runtime error surfaced by the editor. */
  error: (error: Error) => void;
  /** The unsaved-changes flag flipped. `true` after the first edit since the
   *  last load/save, `false` on save / `setContent` / `import`. */
  dirtyChange: (dirty: boolean) => void;
}

export interface CasualSheetsAPI {
  /** Current workbook as an `IWorkbookData` snapshot — the canonical
   *  cross-editor content accessor (doc 38 §4). `null` before the unit is
   *  created (shouldn't happen after `onReady`, but typed defensively). */
  getContent(): IWorkbookData | null;
  /** Replace the workbook with a new snapshot — the canonical cross-editor
   *  content setter (doc 38 §4). Disposes the current unit, mounts `data` as a
   *  fresh one, and clears the dirty flag. */
  setContent(data: IWorkbookData): void;
  /**
   * @deprecated Use {@link getContent} instead. Kept as an alias for one minor.
   */
  getSnapshot(): IWorkbookData | null;
  /**
   * @deprecated Use {@link setContent} instead. Kept as an alias for one minor.
   */
  loadSnapshot(data: IWorkbookData): void;
  /** Parse an `.xlsx` and load it as the active workbook — the format-specific
   *  file import (canonically aliased as {@link import}). Accepts a `File` /
   *  `Blob` (e.g. from an `<input type=file>`), an `ArrayBuffer`, or a
   *  `Uint8Array`. The ExcelJS parser is lazy-loaded as a separate chunk, so
   *  hosts that never import a file don't pay for it. When a `File` is passed,
   *  its name + on-disk size are recorded on the snapshot (surfaced by the
   *  built-in Properties dialog). Resolves to the loaded snapshot. */
  importXlsx(input: ArrayBuffer | Uint8Array | Blob): Promise<IWorkbookData>;
  /** Serialize the current workbook to an `.xlsx` `Blob` — the format-specific
   *  file export (canonically aliased as {@link export}). Covers the core
   *  fidelity (values/formulas, styles, merges, number formats, borders,
   *  hyperlinks, comments, data validation, tables, page setup, named ranges,
   *  VBA passthrough) — everything carried on the snapshot. App-level extras
   *  (chart/pivot/sparkline models) are a power-host concern and aren't included
   *  here. The converter (ExcelJS) is lazy-loaded as a separate chunk. Rejects
   *  if there is no active workbook. */
  exportXlsx(): Promise<Blob>;
  /** Canonical cross-editor file import (doc 38 §4) — an alias of
   *  {@link importXlsx}. */
  import(input: ArrayBuffer | Uint8Array | Blob): Promise<IWorkbookData>;
  /** Canonical cross-editor file export (doc 38 §4) — an alias of
   *  {@link exportXlsx}. */
  export(): Promise<Blob>;
  /** The active selection, or `null` when there is none. */
  getSelection(): RangeRef | null;
  /** Move keyboard focus to the active workbook so subsequent typing / shortcuts
   *  land on the grid (doc 38 §4). No-op when there is no active workbook. */
  focus(): void;
  /** Subscribe to a canonical editor event (doc 38 §3). Returns an unsubscribe
   *  function. `'ready'` is sticky — subscribing after the editor is already
   *  ready invokes the handler immediately. */
  on<K extends keyof CasualSheetsEvents>(name: K, handler: CasualSheetsEvents[K]): () => void;
  /** Remove a previously-registered event handler (doc 38 §3). */
  off<K extends keyof CasualSheetsEvents>(name: K, handler: CasualSheetsEvents[K]): void;
  /** Dispatch a Univer command by id. Resolves to the command's boolean
   *  result. */
  executeCommand(id: string, params?: object): Promise<boolean>;
  /** Replay a sequence of command/mutation steps in order — e.g. a recorded
   *  macro, or a host-authored script. Best-effort: a step that throws is
   *  skipped (the underlying state may have moved on). Resolves to the number
   *  of steps that ran without throwing. */
  executeCommands(steps: CommandRecord[]): Promise<number>;
  /** Undo the last edit — the canonical cross-editor history control (doc 38
   *  §4). Dispatches Univer's `univer.command.undo` on the active unit
   *  (fire-and-forget: no-op when there is nothing to undo). */
  undo(): void;
  /** Redo the last undone edit — the canonical cross-editor history control
   *  (doc 38 §4). Dispatches Univer's `univer.command.redo` (fire-and-forget:
   *  no-op when there is nothing to redo). */
  redo(): void;
  /** Observe the replayable mutation stream so a host can record automations
   *  or build an audit log. Wraps Univer's canonical collab hook
   *  (`onMutationExecutedForCollab`): fires for `CommandType.MUTATION` only —
   *  the deterministic, replayable state changes, not transient command/calc
   *  noise. Pair with `executeCommands` for record→replay. Returns a disposer;
   *  call it to stop observing. */
  onMutation(handler: (record: CommandRecord) => void): () => void;
  /** Imperative light/dark switch — the API equivalent of the reactive
   *  `appearance` prop. Flips Univer's `ThemeService.setDarkMode` (canvas
   *  colours + the `univer-dark` class Univer applies to the document root). */
  setTheme(appearance: 'light' | 'dark'): void;
  /** Switch the active workbook between `'editing'` (fully editable, the
   *  default) and `'viewing'` (read-only). `'viewing'` applies the same
   *  command-veto + `WorkbookEditablePermission` path as {@link applyReadOnly};
   *  `'editing'` disposes it and restores the prior editable state. Idempotent —
   *  re-applying the current mode is a no-op. No-op when there is no active
   *  workbook. */
  setDocumentMode(mode: DocumentMode): void;
  /** The current document mode: `'viewing'` while read-only is applied via
   *  {@link setDocumentMode}, else `'editing'`. */
  getDocumentMode(): DocumentMode;
  /** The FUniver facade — documented escape hatch, NOT covered by semver. */
  univer: FUniver;
}

/**
 * Internal view of the API used by the `<CasualSheets>` wrapper to drive the
 * unified emitter from seams the factory can't see on its own (`ready` timing,
 * the Ctrl/Cmd+S save, boot errors) and to flip the dirty flag from the
 * wrapper's mutation subscription. NOT part of the public, semver-covered
 * surface — hosts use {@link CasualSheetsAPI}.
 * @internal
 */
export interface CasualSheetsAPIInternal extends CasualSheetsAPI {
  /** Emit a canonical event to all `api.on(name, …)` subscribers. */
  emit<K extends keyof CasualSheetsEvents>(
    name: K,
    ...args: Parameters<CasualSheetsEvents[K]>
  ): void;
  /** Flip the unsaved-changes flag; emits `dirtyChange` only on transitions. */
  markDirty(dirty: boolean): void;
  /** Live subscriber count for an event — lets the wrapper skip work (e.g.
   *  serializing a snapshot for `change`) when nothing is listening. */
  listenerCount<K extends keyof CasualSheetsEvents>(name: K): number;
}

/**
 * Build the imperative API over a live FUniver facade. The wrapper holds no
 * state of its own — every call reads the current active workbook, so it stays
 * correct across `setContent` swaps without the host re-acquiring the ref.
 *
 * Returns a {@link CasualSheetsAPIInternal}; the public return type narrows it
 * to {@link CasualSheetsAPI} for hosts, while the `<CasualSheets>` wrapper casts
 * back to reach `emit` / `markDirty`.
 */
export function createCasualSheetsAPI(univerAPI: FUniver): CasualSheetsAPI {
  // Extracted so importXlsx / setContent reuse the exact same swap semantics.
  const swapWorkbook = (data: IWorkbookData) => {
    const current = univerAPI.getActiveWorkbook();
    if (current) univerAPI.disposeUnit(current.getId());
    univerAPI.createWorkbook(data);
  };

  // Unified event emitter (doc 38 §3). Drives api.on/off; the factory wires the
  // Univer-derived events (`selectionChange`) itself, and the <CasualSheets>
  // wrapper drives the seams it owns (`ready`/`change`/`save`/`error`) via emit.
  const emitter = createEmitter<CasualSheetsEvents>();
  // `ready` is sticky: a late `api.on('ready', …)` fires immediately. We remember
  // the API instance passed to the first `emit('ready', …)`.
  let readyFired = false;
  let readyApi: CasualSheetsAPI | null = null;

  // Dirty tracking: flips true on the first buffered mutation after a load/save,
  // false on setContent/import/save. Emits `dirtyChange` on transitions only.
  let dirty = false;
  const markDirty = (next: boolean) => {
    if (next === dirty) return;
    dirty = next;
    emitter.emit('dirtyChange', next);
  };

  // Document-mode state: `'editing'` by default. When `'viewing'`, we hold the
  // disposer returned by applyReadOnly so switching back to `'editing'` restores
  // the prior editable state instead of reinventing the permission path.
  let documentMode: DocumentMode = 'editing';
  let readOnlyDisposer: (() => void) | null = null;

  const getContent = (): IWorkbookData | null => univerAPI.getActiveWorkbook()?.save() ?? null;

  const setContent = (data: IWorkbookData): void => {
    swapWorkbook(data);
    // A fresh snapshot is a clean buffer until the user edits it.
    markDirty(false);
  };

  const importXlsx = async (input: ArrayBuffer | Uint8Array | Blob): Promise<IWorkbookData> => {
    // Normalise to ArrayBuffer. Blob/File expose arrayBuffer(); a Uint8Array
    // view is sliced to its exact window so we don't hand the parser a larger
    // backing buffer.
    let buffer: ArrayBuffer;
    if (input instanceof ArrayBuffer) {
      buffer = input;
    } else if (input instanceof Uint8Array) {
      // `.slice` of a (possibly SharedArrayBuffer-backed) view; uploads are
      // never shared, so narrow to ArrayBuffer for the parser.
      buffer = input.buffer.slice(
        input.byteOffset,
        input.byteOffset + input.byteLength,
      ) as ArrayBuffer;
    } else {
      buffer = await input.arrayBuffer();
    }
    // Bare subpath import → separate chunk (see file header + tsup external).
    const { xlsxToWorkbookData } = await import('@casualoffice/sheets/xlsx');
    const data = await xlsxToWorkbookData(buffer);
    // A File carries the original name + size; surface them on the snapshot so
    // the built-in Properties dialog shows the real file (not the snapshot).
    if (typeof Blob !== 'undefined' && input instanceof Blob && 'name' in input) {
      const file = input as File;
      data.name = file.name.replace(/\.(xlsx|xlsm)$/i, '') || data.name;
      data.custom = { ...data.custom, sourceBytes: file.size, sourceName: file.name };
    }
    swapWorkbook(data);
    markDirty(false);
    return data;
  };

  const exportXlsx = async (): Promise<Blob> => {
    const snap = univerAPI.getActiveWorkbook()?.save();
    if (!snap) throw new Error('exportXlsx: no active workbook to export');
    // Bare subpath import → separate chunk (see file header + tsup external).
    const { workbookDataToXlsx } = await import('@casualoffice/sheets/xlsx');
    return workbookDataToXlsx(snap as IWorkbookData);
  };

  const getSelection = (): RangeRef | null => {
    const wb = univerAPI.getActiveWorkbook();
    const range = wb?.getActiveRange();
    if (!wb || !range) return null;
    return {
      unitId: wb.getId(),
      sheetId: wb.getActiveSheet().getSheetId(),
      range: range.getRange(),
    };
  };

  const focus = (): void => {
    const wb = univerAPI.getActiveWorkbook();
    if (!wb) return;
    try {
      const injector = (univerAPI as unknown as { _injector?: { get(t: unknown): unknown } })
        ._injector;
      const instanceService = injector?.get(IUniverInstanceService) as
        | { focusUnit(id: string | null): void }
        | undefined;
      // The same focus path FWorkbook.undo/redo use before dispatching, so the
      // grid takes keyboard focus for subsequent typing / shortcuts.
      instanceService?.focusUnit(wb.getId());
    } catch {
      /* focus is best-effort — no facade/injector, or unit already gone */
    }
  };

  // Sticky `on`: a late `on('ready', …)` fires immediately with the live API.
  const on = <K extends keyof CasualSheetsEvents>(
    name: K,
    handler: CasualSheetsEvents[K],
  ): (() => void) => {
    const unsubscribe = emitter.on(name, handler);
    if (name === 'ready' && readyFired && readyApi) {
      (handler as CasualSheetsEvents['ready'])(readyApi);
    }
    return unsubscribe;
  };

  const emit = <K extends keyof CasualSheetsEvents>(
    name: K,
    ...args: Parameters<CasualSheetsEvents[K]>
  ): void => {
    if (name === 'ready') {
      readyFired = true;
      readyApi = args[0] as CasualSheetsAPI;
    }
    emitter.emit(name, ...args);
  };

  // Canvas-driven selection → `selectionChange`. Both the internal
  // `SelectionChanged` (programmatic + user) and `SelectionMoveEnd` (fires
  // reliably on drag-release) feed it, mirroring the chrome FormulaBar's picker
  // wiring. Guarded by listener count so bare hosts never pay for `getSelection`.
  const emitSelection = () => {
    if (emitter.listenerCount('selectionChange') === 0) return;
    emitter.emit('selectionChange', getSelection());
  };
  // `SelectionChanged` / `SelectionMoveEnd` are contributed by the sheets-ui
  // facade (`@univerjs/sheets-ui/facade`), which the HOST registers (apps/web's
  // `facade.ts`). We reach them by name through a loose cast rather than a static
  // import: the vendored fork maps `@univerjs/sheets-ui/facade` to raw `.ts`
  // source, so importing it here would drag ~36 pre-existing fork type errors
  // into this package's typecheck. When the facade is present the events fire;
  // otherwise this stays a harmless no-op and the other events are unaffected.
  const evented = univerAPI as unknown as {
    Event?: Record<string, string | undefined>;
    addEvent(event: string, cb: () => void): { dispose(): void };
  };
  try {
    const changed = evented.Event?.SelectionChanged;
    const moveEnd = evented.Event?.SelectionMoveEnd;
    if (changed) evented.addEvent(changed, emitSelection);
    if (moveEnd) evented.addEvent(moveEnd, emitSelection);
  } catch {
    /* facade not loaded — selection events unavailable; other events unaffected */
  }

  const api: CasualSheetsAPIInternal = {
    univer: univerAPI,

    getContent,
    setContent,
    // Deprecated aliases (doc 38 §4) — kept working for one minor.
    getSnapshot: getContent,
    loadSnapshot: setContent,

    importXlsx,
    exportXlsx,
    // Canonical cross-editor aliases (doc 38 §4).
    import: importXlsx,
    export: exportXlsx,

    getSelection,
    focus,

    on,
    off: emitter.off,
    emit,
    markDirty,
    listenerCount: emitter.listenerCount,

    executeCommand(id, params) {
      return univerAPI.executeCommand(id, params) as Promise<boolean>;
    },

    executeCommands(steps) {
      return runSteps((id, params) => univerAPI.executeCommand(id, params), steps);
    },

    // History controls (doc 38 §4). Dispatched through the same command path the
    // built-in chrome (Toolbar/MenuBar) and the app shell already use, so undo/
    // redo, collab replay, and dirty tracking all stay consistent. Void the
    // promise — the canonical handle types these as fire-and-forget.
    undo() {
      void univerAPI.executeCommand('univer.command.undo');
    },

    redo() {
      void univerAPI.executeCommand('univer.command.redo');
    },

    onMutation(handler) {
      const injector = (univerAPI as unknown as { _injector?: { get(t: unknown): unknown } })
        ._injector;
      const cmdSvc = injector?.get(ICommandService) as MutationEmitter | undefined;
      return attachMutationObserver(cmdSvc, handler);
    },

    setTheme(appearance) {
      const dark = appearance === 'dark';
      const injector = (univerAPI as unknown as { _injector?: { get(t: unknown): unknown } })
        ._injector;
      const themeService = injector?.get(ThemeService) as
        | { setDarkMode(b: boolean): void; darkMode: boolean }
        | undefined;
      if (themeService && themeService.darkMode !== dark) themeService.setDarkMode(dark);
    },

    setDocumentMode(mode) {
      if (mode === documentMode) return;
      if (mode === 'viewing') {
        const wb = univerAPI.getActiveWorkbook();
        if (!wb) return; // no workbook yet — stay 'editing', caller can retry
        readOnlyDisposer = applyReadOnly(univerAPI, wb.getId());
        documentMode = 'viewing';
      } else {
        readOnlyDisposer?.();
        readOnlyDisposer = null;
        documentMode = 'editing';
      }
    },

    getDocumentMode() {
      return documentMode;
    },
  };

  return api;
}
