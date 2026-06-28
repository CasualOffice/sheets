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

import { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { useUniverAPI } from '../use-univer';
import { useUI } from '../use-ui';
import { getHeaderGutter, getUniverHost, getUniverMainCanvas } from '../univer-dom';

/**
 * Excel-style "Show Formulas" overlay (Ctrl+`). When toggled on, every
 * cell with a formula is painted with an opaque text panel showing the
 * formula source instead of the computed value.
 *
 * Implementation choice: a DOM overlay sitting above the Univer canvas
 * via a portal. The alternative — swapping `v` ↔ `f` via mutations —
 * would muddy the undo stack, the autosave snapshot, and the collab
 * op-log, and would lose data if the user edited cells while the mode
 * was active. The overlay path is non-destructive and reversible.
 *
 * Tradeoffs accepted in v1:
 *   - Overlay text uses a fixed monospace font, not the cell's actual
 *     style. Excel does the same in Show-Formulas mode.
 *   - Overlay covers the cell content but not the gridlines — the
 *     opaque background sits inside the cell rect.
 *   - Only the active sheet is overlaid. Switching sheets re-runs the
 *     scan automatically because the effect depends on `meta.revision`
 *     via the cell-data lookup below.
 */

type FormulaCell = {
  row: number;
  col: number;
  f: string;
  /** Cell-rect in host coords + height for the overlay positioning. */
  left: number;
  top: number;
  width: number;
  height: number;
};

export function ShowFormulasLayer() {
  const api = useUniverAPI();
  const ui = useUI();
  const [cells, setCells] = useState<FormulaCell[]>([]);
  const hostRef = useRef<HTMLElement | null>(null);
  const cellsRef = useRef<FormulaCell[]>(cells);
  cellsRef.current = cells;

  useEffect(() => {
    hostRef.current = getUniverHost();
  }, []);

  useEffect(() => {
    if (!api || !ui.showFormulas) {
      if (cellsRef.current.length) setCells([]);
      return;
    }
    let raf = 0;
    const tick = () => {
      try {
        recompute();
      } catch (err) {
        console.debug('[show-formulas] recompute threw', err);
      }
      raf = requestAnimationFrame(tick);
    };
    const recompute = () => {
      const host = hostRef.current ?? getUniverHost();
      if (!host) return;
      const canvas = getUniverMainCanvas(host);
      if (!canvas) return;
      const hostRect = host.getBoundingClientRect();
      const canvasRect = canvas.getBoundingClientRect();
      const dx = canvasRect.left - hostRect.left;
      const dy = canvasRect.top - hostRect.top;
      const gutter = getHeaderGutter(api);

      const wb = api.getActiveWorkbook();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const activeSheet = wb?.getActiveSheet() as any;
      if (!wb || !activeSheet) return;

      // Pull formula cells from the workbook snapshot — `getSheet()`
      // exposes the underlying Worksheet, whose `cellMatrix` has the
      // sparse cellData (row → col → ICellData). Iterating only the
      // cells that actually have entries is far cheaper than scanning
      // the visible rectangle row-by-row.
      const ws = activeSheet.getSheet?.();
      const formulaCells: Array<{ row: number; col: number; f: string }> = [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const cellMatrix = (ws as any)?.getCellMatrix?.() as
        | {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            forValue: (cb: (row: number, col: number, cell: any) => void) => void;
          }
        | undefined;
      if (cellMatrix?.forValue) {
        cellMatrix.forValue((row, col, cell) => {
          if (cell?.f && typeof cell.f === 'string') {
            formulaCells.push({ row, col, f: String(cell.f) });
          }
        });
      } else {
        // Fallback: scan the active selection's data range via the
        // facade. Older Univer builds may not expose cellMatrix.
        const last = activeSheet.getLastRowWithContent?.() ?? 0;
        const lastCol = activeSheet.getLastColumnWithContent?.() ?? 0;
        for (let r = 0; r <= last; r += 1) {
          for (let c = 0; c <= lastCol; c += 1) {
            const data = activeSheet.getRange(r, c).getCellData();
            if (data?.f && typeof data.f === 'string') {
              formulaCells.push({ row: r, col: c, f: String(data.f) });
            }
          }
        }
      }

      // Scroll + zoom maths mirror SparklineLayer.
      let sx = 0;
      let sy = 0;
      try {
        const scrollState = activeSheet.getScrollState?.() as
          | { sheetViewStartRow?: number; sheetViewStartColumn?: number; offsetX?: number; offsetY?: number }
          | undefined;
        if (scrollState) {
          const r = scrollState.sheetViewStartRow ?? 0;
          const c = scrollState.sheetViewStartColumn ?? 0;
          const topLeft = activeSheet.getRange(r, c).getCellRect();
          if (topLeft) {
            sx = topLeft.left + (scrollState.offsetX ?? 0);
            sy = topLeft.top + (scrollState.offsetY ?? 0);
          }
        }
      } catch {
        /* skeleton not ready — leave at 0 this frame */
      }
      let zoom = 1;
      try {
        const z =
          (activeSheet?._worksheet?.getZoomRatio?.() as number | undefined) ??
          (activeSheet?.getZoomRatio?.() as number | undefined);
        if (typeof z === 'number' && z > 0) zoom = z;
      } catch {
        /* default 1 */
      }

      const out: FormulaCell[] = [];
      for (const f of formulaCells) {
        try {
          const rect = activeSheet.getRange(f.row, f.col).getCellRect();
          if (!rect) continue;
          const left = (rect.left - sx) * zoom + dx + gutter.rowHeaderWidth;
          const top = (rect.top - sy) * zoom + dy + gutter.columnHeaderHeight;
          const right = (rect.right - sx) * zoom + dx + gutter.rowHeaderWidth;
          const bottom = (rect.bottom - sy) * zoom + dy + gutter.columnHeaderHeight;
          // Clip off-screen so the overlay doesn't sit on top of
          // headers or extend into the column-letter gutter.
          if (right < dx || bottom < dy) continue;
          if (left > dx + canvasRect.width || top > dy + canvasRect.height) continue;
          out.push({
            row: f.row,
            col: f.col,
            f: f.f.startsWith('=') ? f.f : `=${f.f}`,
            left,
            top,
            width: Math.max(8, right - left),
            height: Math.max(8, bottom - top),
          });
        } catch {
          /* getCellRect can throw mid-resize */
        }
      }
      if (cellsEqual(out, cellsRef.current)) return;
      setCells(out);
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
  }, [api, ui.showFormulas]);

  const host = hostRef.current ?? getUniverHost();
  if (!host || !ui.showFormulas) return null;
  if (cells.length === 0) return null;

  return createPortal(
    <div className="show-formulas-layer" data-testid="show-formulas-layer" aria-hidden="true">
      {cells.map((c) => (
        <div
          key={`${c.row}:${c.col}`}
          className="show-formulas-cell"
          data-testid="show-formulas-cell"
          style={{ left: c.left, top: c.top, width: c.width, height: c.height }}
          title={c.f}
        >
          {c.f}
        </div>
      ))}
    </div>,
    host,
  );
}

function cellsEqual(a: FormulaCell[], b: FormulaCell[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) {
    const x = a[i];
    const y = b[i];
    if (
      x.row !== y.row ||
      x.col !== y.col ||
      x.left !== y.left ||
      x.top !== y.top ||
      x.width !== y.width ||
      x.height !== y.height ||
      x.f !== y.f
    )
      return false;
  }
  return true;
}
