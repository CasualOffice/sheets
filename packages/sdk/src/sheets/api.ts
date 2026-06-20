/**
 * CasualSheetsAPI — the imperative ref handed to a host via `<CasualSheets onReady>`.
 *
 * This is the SDK's stable integration surface (Excalidraw's model: props +
 * imperative ref). Hosts drive the editor through these methods rather than
 * reaching into Univer directly; `api.univer` is the documented escape hatch and
 * is explicitly NOT covered by semver — everything else here is.
 *
 * Scope of THIS batch (SDK restructure, Phase 1 step 3):
 *   getSnapshot / loadSnapshot / getSelection / executeCommand / univer
 *
 * Deferred to later, clearly-scoped batches (kept off the type until they work,
 * so the surface never advertises a method that throws):
 *   - importXlsx / exportXlsx — the xlsx I/O batch. importXlsx must NOT be a
 *     plain `import('../xlsx')` here: the main tsup config is `splitting:false`,
 *     so a dynamic import gets inlined and balloons the editor entry from ~11KB
 *     to ~200KB of parser code for hosts that never open a file. The xlsx-I/O
 *     batch wires it as its own chunk (and lifts the export converter out of
 *     apps/web — the SDK xlsx module is import-only today).
 *   - attachCollab — belongs to the storage/collab adapter phase (Phase 2);
 *     the editor ships collab-unaware until then.
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
  return {
    univer: univerAPI,

    getSnapshot() {
      return univerAPI.getActiveWorkbook()?.save() ?? null;
    },

    loadSnapshot(data) {
      const current = univerAPI.getActiveWorkbook();
      if (current) univerAPI.disposeUnit(current.getId());
      univerAPI.createWorkbook(data);
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
