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
import { getHeaderGutter, getUniverHost, getUniverMainCanvas } from '../univer-dom';
import { formulaReferences, precedentsOf, type CellRect } from './trace-model';

/**
 * Formula auditing arrows (Excel's Trace Precedents / Trace Dependents). A
 * grid overlay — like SparklineLayer / ShowFormulasLayer — that draws an arrow
 * from each precedent cell to the active cell (or from the active cell to each
 * dependent). Triggered by menu CustomEvents; arrows track scroll/zoom via a
 * rAF loop and clear on sheet switch or `Remove Arrows`.
 *
 * The reference parsing is pure + unit-tested (trace-model.ts); this resolves
 * pixel geometry the same way the other overlays do.
 */

const DEPENDENT_SCAN_CAP = 5000; // bound the used-range scan for dependents

type Arrow = { from: CellRect; to: CellRect };
type Trace = { sheetId: string; arrows: Arrow[] };
type PixelArrow = { x1: number; y1: number; x2: number; y2: number };

export function TracePrecedentsLayer() {
  const api = useUniverAPI();
  const [trace, setTrace] = useState<Trace | null>(null);
  const [pixels, setPixels] = useState<PixelArrow[]>([]);
  const hostRef = useRef<HTMLElement | null>(null);
  const traceRef = useRef<Trace | null>(null);
  const pixelsRef = useRef<PixelArrow[]>(pixels);
  traceRef.current = trace;
  pixelsRef.current = pixels;

  useEffect(() => {
    hostRef.current = getUniverHost();
  }, []);

  // Menu-triggered: compute precedent / dependent arrows for the active cell.
  useEffect(() => {
    if (!api) return;
    const activeCellRect = (): { sheetId: string; sheetName: string; cell: CellRect } | null => {
      const wb = api.getActiveWorkbook();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws = wb?.getActiveSheet() as any;
      const r = ws?.getActiveRange?.();
      if (!ws || !r) return null;
      return {
        sheetId: ws.getSheetId?.() ?? '',
        sheetName: ws.getSheetName?.() ?? '',
        cell: {
          startRow: r.getRow(),
          startCol: r.getColumn(),
          endRow: r.getRow(),
          endCol: r.getColumn(),
        },
      };
    };

    const onPrecedents = () => {
      const active = activeCellRect();
      if (!active) return;
      const wb = api.getActiveWorkbook();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws = wb?.getActiveSheet() as any;
      const formula =
        ws?.getRange?.(active.cell.startRow, active.cell.startCol)?.getFormula?.() ?? '';
      if (!formula.startsWith('=')) {
        setTrace(null);
        return;
      }
      const arrows: Arrow[] = precedentsOf(formula)
        // Same-sheet precedents only (cross-sheet arrows would leave the view).
        .filter((t) => t.sheetName == null || t.sheetName === active.sheetName)
        .map((t) => ({ from: t.rect, to: active.cell }));
      setTrace(arrows.length ? { sheetId: active.sheetId, arrows } : null);
    };

    const onDependents = () => {
      const active = activeCellRect();
      if (!active) return;
      const wb = api.getActiveWorkbook();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws = wb?.getActiveSheet() as any;
      // Scan the used range for formulas referencing the active cell.
      let used: { endRow?: number; endColumn?: number } | null = null;
      try {
        used = ws?.getDataRange?.()?.getRange?.() ?? null;
      } catch {
        used = null;
      }
      const maxRow = Math.min(used?.endRow ?? 200, 5000);
      const maxCol = Math.min(used?.endColumn ?? 50, 200);
      const arrows: Arrow[] = [];
      let scanned = 0;
      for (let r = 0; r <= maxRow && scanned < DEPENDENT_SCAN_CAP; r++) {
        for (let c = 0; c <= maxCol && scanned < DEPENDENT_SCAN_CAP; c++) {
          scanned++;
          if (r === active.cell.startRow && c === active.cell.startCol) continue;
          let f = '';
          try {
            f = ws?.getRange?.(r, c)?.getFormula?.() ?? '';
          } catch {
            continue;
          }
          if (!f.startsWith('=')) continue;
          if (formulaReferences(f, active.sheetName, active.cell.startRow, active.cell.startCol)) {
            arrows.push({
              from: active.cell,
              to: { startRow: r, startCol: c, endRow: r, endCol: c },
            });
          }
        }
      }
      setTrace(arrows.length ? { sheetId: active.sheetId, arrows } : null);
    };

    const onClear = () => setTrace(null);

    document.addEventListener('casual-trace-precedents', onPrecedents);
    document.addEventListener('casual-trace-dependents', onDependents);
    document.addEventListener('casual-trace-clear', onClear);
    return () => {
      document.removeEventListener('casual-trace-precedents', onPrecedents);
      document.removeEventListener('casual-trace-dependents', onDependents);
      document.removeEventListener('casual-trace-clear', onClear);
    };
  }, [api]);

  // Clear arrows when the active sheet changes (they'd point at the wrong grid).
  useEffect(() => {
    if (!api) return;
    const disp = api.addEvent(api.Event.ActiveSheetChanged, () => setTrace(null));
    return () => disp.dispose();
  }, [api]);

  // rAF: resolve pixel geometry, tracking scroll / zoom (same approach as the
  // other grid overlays).
  useEffect(() => {
    if (!api || !trace) {
      if (pixelsRef.current.length) setPixels([]);
      return;
    }
    let raf = 0;
    const tick = () => {
      try {
        recompute();
      } catch {
        /* skeleton mid-resize */
      }
      raf = requestAnimationFrame(tick);
    };
    const recompute = () => {
      const t = traceRef.current;
      if (!t) return;
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
      const ws = wb?.getActiveSheet() as any;
      if (!ws || ws.getSheetId?.() !== t.sheetId) {
        if (pixelsRef.current.length) setPixels([]);
        return;
      }

      let sx = 0;
      let sy = 0;
      try {
        const ss = ws.getScrollState?.() as
          | {
              sheetViewStartRow?: number;
              sheetViewStartColumn?: number;
              offsetX?: number;
              offsetY?: number;
            }
          | undefined;
        if (ss) {
          const tl = ws
            .getRange(ss.sheetViewStartRow ?? 0, ss.sheetViewStartColumn ?? 0)
            .getCellRect();
          if (tl) {
            sx = tl.left + (ss.offsetX ?? 0);
            sy = tl.top + (ss.offsetY ?? 0);
          }
        }
      } catch {
        /* not ready */
      }
      let zoom = 1;
      try {
        const z = ws._worksheet?.getZoomRatio?.() ?? ws.getZoomRatio?.();
        if (typeof z === 'number' && z > 0) zoom = z;
      } catch {
        /* default */
      }

      const center = (rect: CellRect): { x: number; y: number } | null => {
        const a = ws.getRange(rect.startRow, rect.startCol).getCellRect();
        const b = ws.getRange(rect.endRow, rect.endCol).getCellRect();
        if (!a || !b) return null;
        const left = (a.left - sx) * zoom + dx + gutter.rowHeaderWidth;
        const top = (a.top - sy) * zoom + dy + gutter.columnHeaderHeight;
        const right = (b.right - sx) * zoom + dx + gutter.rowHeaderWidth;
        const bottom = (b.bottom - sy) * zoom + dy + gutter.columnHeaderHeight;
        return { x: (left + right) / 2, y: (top + bottom) / 2 };
      };

      const out: PixelArrow[] = [];
      for (const arrow of t.arrows) {
        const from = center(arrow.from);
        const to = center(arrow.to);
        if (!from || !to) continue;
        out.push({ x1: from.x, y1: from.y, x2: to.x, y2: to.y });
      }
      if (pixelArrowsEqual(out, pixelsRef.current)) return;
      setPixels(out);
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
  }, [api, trace]);

  const host = hostRef.current ?? getUniverHost();
  if (!host || pixels.length === 0) return null;

  return createPortal(
    <svg className="trace-layer" data-testid="trace-layer" aria-hidden="true">
      <defs>
        <marker
          id="trace-arrowhead"
          markerWidth="8"
          markerHeight="8"
          refX="6"
          refY="3"
          orient="auto"
        >
          <path d="M0,0 L6,3 L0,6 Z" fill="var(--color-accent, #107c41)" />
        </marker>
      </defs>
      {pixels.map((p, i) => (
        <line
          key={i}
          data-testid="trace-arrow"
          x1={p.x1}
          y1={p.y1}
          x2={p.x2}
          y2={p.y2}
          stroke="var(--color-accent, #107c41)"
          strokeWidth={1.5}
          markerEnd="url(#trace-arrowhead)"
        />
      ))}
    </svg>,
    host,
  );
}

function pixelArrowsEqual(a: PixelArrow[], b: PixelArrow[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i].x1 !== b[i].x1 || a[i].y1 !== b[i].y1 || a[i].x2 !== b[i].x2 || a[i].y2 !== b[i].y2) {
      return false;
    }
  }
  return true;
}
