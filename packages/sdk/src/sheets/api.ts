/**
 * CasualSheetsAPI — the imperative ref handed to a host via `<CasualSheets onReady>`.
 *
 * This is the SDK's stable integration surface (Excalidraw's model: props +
 * imperative ref). Hosts drive the editor through these methods rather than
 * reaching into Univer directly; `api.univer` is the documented escape hatch and
 * is explicitly NOT covered by semver — everything else here is.
 *
 * Surface:
 *   getSnapshot / loadSnapshot / getSelection / executeCommand / setTheme /
 *   importXlsx / exportXlsx / univer
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
import { ThemeService } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import type { IRange, IWorkbookData } from '@univerjs/core';

/** The active selection, as a sheet-scoped range. */
export interface RangeRef {
  /** Workbook unit id the selection belongs to. */
  unitId: string;
  /** Worksheet (sub-unit) id the selection belongs to. */
  sheetId: string;
  /** `{ startRow, startColumn, endRow, endColumn }`. */
  range: IRange;
}

export interface CasualSheetsAPI {
  /** Current workbook as an `IWorkbookData` snapshot. `null` before the unit
   *  is created (shouldn't happen after `onReady`, but typed defensively). */
  getSnapshot(): IWorkbookData | null;
  /** Replace the workbook with a new snapshot. Disposes the current unit and
   *  mounts `data` as a fresh one. */
  loadSnapshot(data: IWorkbookData): void;
  /** Parse an `.xlsx` and load it as the active workbook. Accepts a `File` /
   *  `Blob` (e.g. from an `<input type=file>`), an `ArrayBuffer`, or a
   *  `Uint8Array`. The ExcelJS parser is lazy-loaded as a separate chunk, so
   *  hosts that never import a file don't pay for it. When a `File` is passed,
   *  its name + on-disk size are recorded on the snapshot (surfaced by the
   *  built-in Properties dialog). Resolves to the loaded snapshot. */
  importXlsx(input: ArrayBuffer | Uint8Array | Blob): Promise<IWorkbookData>;
  /** Serialize the current workbook to an `.xlsx` `Blob`. Covers the core
   *  fidelity (values/formulas, styles, merges, number formats, borders,
   *  hyperlinks, comments, data validation, tables, page setup, named ranges,
   *  VBA passthrough) — everything carried on the snapshot. App-level extras
   *  (chart/pivot/sparkline models) are a power-host concern and aren't included
   *  here. The converter (ExcelJS) is lazy-loaded as a separate chunk. Rejects
   *  if there is no active workbook. */
  exportXlsx(): Promise<Blob>;
  /** The active selection, or `null` when there is none. */
  getSelection(): RangeRef | null;
  /** Dispatch a Univer command by id. Resolves to the command's boolean
   *  result. */
  executeCommand(id: string, params?: object): Promise<boolean>;
  /** Imperative light/dark switch — the API equivalent of the reactive
   *  `appearance` prop. Flips Univer's `ThemeService.setDarkMode` (canvas
   *  colours + the `univer-dark` class Univer applies to the document root). */
  setTheme(appearance: 'light' | 'dark'): void;
  /** The FUniver facade — documented escape hatch, NOT covered by semver. */
  univer: FUniver;
}

/**
 * Build the imperative API over a live FUniver facade. The wrapper holds no
 * state of its own — every call reads the current active workbook, so it stays
 * correct across `loadSnapshot` swaps without the host re-acquiring the ref.
 */
export function createCasualSheetsAPI(univerAPI: FUniver): CasualSheetsAPI {
  // Extracted so importXlsx can reuse the exact same swap semantics.
  const loadSnapshot = (data: IWorkbookData) => {
    const current = univerAPI.getActiveWorkbook();
    if (current) univerAPI.disposeUnit(current.getId());
    univerAPI.createWorkbook(data);
  };

  return {
    univer: univerAPI,

    getSnapshot() {
      return univerAPI.getActiveWorkbook()?.save() ?? null;
    },

    loadSnapshot,

    async importXlsx(input) {
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
      loadSnapshot(data);
      return data;
    },

    async exportXlsx() {
      const snap = univerAPI.getActiveWorkbook()?.save();
      if (!snap) throw new Error('exportXlsx: no active workbook to export');
      // Bare subpath import → separate chunk (see file header + tsup external).
      const { workbookDataToXlsx } = await import('@casualoffice/sheets/xlsx');
      return workbookDataToXlsx(snap as IWorkbookData);
    },

    getSelection() {
      const wb = univerAPI.getActiveWorkbook();
      const range = wb?.getActiveRange();
      if (!wb || !range) return null;
      return {
        unitId: wb.getId(),
        sheetId: wb.getActiveSheet().getSheetId(),
        range: range.getRange(),
      };
    },

    executeCommand(id, params) {
      return univerAPI.executeCommand(id, params) as Promise<boolean>;
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
  };
}
