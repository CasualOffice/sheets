import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { readOutlineFromSnapshot } from './resources';
import type { OutlineGroup, OutlineState, SheetOutline } from './types';

type Axis = 'rows' | 'cols';

type Ctx = {
  state: OutlineState;
  getSheet: (sheetId: string) => SheetOutline;
  addGroup: (sheetId: string, axis: Axis, start: number, end: number) => OutlineGroup | null;
  removeGroup: (sheetId: string, axis: Axis, id: string) => void;
  setCollapsed: (sheetId: string, axis: Axis, id: string, collapsed: boolean) => void;
};

const EMPTY_SHEET: SheetOutline = { rows: [], cols: [] };

const OutlineContext = createContext<Ctx | null>(null);

/** Hook — consumes the OutlineContext or throws. */
export function useOutline(): Ctx {
  const ctx = useContext(OutlineContext);
  if (!ctx) throw new Error('useOutline must be used inside <OutlineProvider>');
  return ctx;
}

type ProviderProps = { children: ReactNode };

/**
 * Holds the in-memory outline state, mirrored to `data.resources` at save
 * time (see `xlsx/export.ts`) and re-hydrated when the active workbook
 * changes. State is keyed by sheet id; toggling a group dispatches Univer's
 * hide / show row/col commands so the rows actually disappear visually.
 */
export function OutlineProvider({ children }: ProviderProps) {
  const api = useUniverAPI();
  const { meta, snapshotRef } = useWorkbook();
  const [state, setState] = useState<OutlineState>(() =>
    snapshotRef.current ? readOutlineFromSnapshot(snapshotRef.current) : {},
  );

  // Re-hydrate every time the workbook is replaced (Open / New). The
  // snapshot itself isn't in React state (Stage 3 memory win) — it
  // lives on `snapshotRef.current` only for the brief window between
  // replaceWorkbook and the post-render flush, which is exactly the
  // window we need to read it from here.
  const lastRevisionRef = useRef(meta.revision);
  useEffect(() => {
    if (lastRevisionRef.current === meta.revision) return;
    lastRevisionRef.current = meta.revision;
    const snap = snapshotRef.current;
    if (!snap) {
      // Ref already cleared (revision bump arrived from a stale tab
      // path) — fall back to whatever Univer currently has. wb.save()
      // is a deep clone so we only do this on rare misses.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb = api?.getActiveWorkbook?.() as any;
      const fresh = wb?.save?.();
      if (fresh) setState(readOutlineFromSnapshot(fresh));
      return;
    }
    setState(readOutlineFromSnapshot(snap));
    // snapshotRef is a stable ref object, never re-created — safe to
    // exclude from deps.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [meta.revision, api]);

  const getSheet = useCallback(
    (sheetId: string) => state[sheetId] ?? EMPTY_SHEET,
    [state],
  );

  const addGroup = useCallback(
    (sheetId: string, axis: Axis, start: number, end: number) => {
      if (end < start) [start, end] = [end, start];
      const id = `g-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
      const group: OutlineGroup = { id, start, end, collapsed: false };
      let inserted = false;
      setState((prev) => {
        const sheet = prev[sheetId] ?? EMPTY_SHEET;
        const groups = sheet[axis];
        // Single-level: reject if the new range overlaps any existing group on
        // the same axis. The caller can show a toast; we just refuse.
        const overlaps = groups.some((g) => !(end < g.start || start > g.end));
        if (overlaps) return prev;
        inserted = true;
        return {
          ...prev,
          [sheetId]: { ...sheet, [axis]: [...groups, group] },
        };
      });
      return inserted ? group : null;
    },
    [],
  );

  const removeGroup = useCallback(
    (sheetId: string, axis: Axis, id: string) => {
      // If the group was collapsed, show its range so the user can see the
      // rows/cols after un-grouping.
      const sheet = state[sheetId];
      const target = sheet?.[axis].find((g) => g.id === id);
      if (api && target?.collapsed) {
        void dispatchVisibility(api, sheetId, axis, target.start, target.end, true);
      }
      setState((prev) => {
        const cur = prev[sheetId];
        if (!cur) return prev;
        return {
          ...prev,
          [sheetId]: { ...cur, [axis]: cur[axis].filter((g) => g.id !== id) },
        };
      });
    },
    [api, state],
  );

  const setCollapsed = useCallback(
    (sheetId: string, axis: Axis, id: string, collapsed: boolean) => {
      const sheet = state[sheetId];
      const target = sheet?.[axis].find((g) => g.id === id);
      if (!target || target.collapsed === collapsed) return;
      if (api) {
        void dispatchVisibility(api, sheetId, axis, target.start, target.end, !collapsed);
      }
      setState((prev) => {
        const cur = prev[sheetId];
        if (!cur) return prev;
        return {
          ...prev,
          [sheetId]: {
            ...cur,
            [axis]: cur[axis].map((g) => (g.id === id ? { ...g, collapsed } : g)),
          },
        };
      });
    },
    [api, state],
  );

  const value = useMemo<Ctx>(
    () => ({ state, getSheet, addGroup, removeGroup, setCollapsed }),
    [state, getSheet, addGroup, removeGroup, setCollapsed],
  );

  return <OutlineContext.Provider value={value}>{children}</OutlineContext.Provider>;
}

/**
 * Dispatch Univer's hide / show row / col command for the given range. The
 * unused axis is clamped to the actual sheet dimensions — earlier we used
 * Number.MAX_SAFE_INTEGER as a "spans the sheet" placeholder, but the
 * set-rows-hidden mutation appears to iterate row/col bounds internally and
 * hangs hard on a 2^53 range (CI: 30s teardown timeout, context unrecoverable).
 */
async function dispatchVisibility(
  api: FUniver,
  subUnitId: string,
  axis: Axis,
  start: number,
  end: number,
  show: boolean,
): Promise<void> {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  const unitId = wb.getId();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const sheet = (wb as any).getSheetBySheetId?.(subUnitId) ?? wb.getActiveSheet();
  // Facade method names vary slightly across versions — fall back to sensible
  // defaults so we never hand Univer a NaN.
  const maxRow =
    Number(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (sheet as any)?.getMaxRows?.() ?? (sheet as any)?.getRowCount?.() ?? 1024,
    ) - 1;
  const maxCol =
    Number(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (sheet as any)?.getMaxColumns?.() ?? (sheet as any)?.getColumnCount?.() ?? 128,
    ) - 1;
  const range =
    axis === 'rows'
      ? { startRow: start, endRow: end, startColumn: 0, endColumn: Math.max(0, maxCol) }
      : { startRow: 0, endRow: Math.max(0, maxRow), startColumn: start, endColumn: end };
  const cmd =
    axis === 'rows'
      ? (show ? 'sheet.command.set-specific-rows-visible' : 'sheet.command.set-rows-hidden')
      : (show ? 'sheet.command.set-col-visible-on-cols' : 'sheet.command.set-col-hidden');
  try {
    await api.executeCommand(cmd, { unitId, subUnitId, ranges: [range] });
  } catch (err) {
    console.warn('[outline]', cmd, 'failed', err);
  }
}
