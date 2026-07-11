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
import { getHeaderGutter, getUniverHost, getUniverMainCanvas } from './univer-dom';
import { useCharts } from './charts-context';
import { ChartOverlay } from './ChartOverlay';
import { ChartContextMenu } from './ChartContextMenu';

/**
 * Renders every chart in the store. Same anchoring strategy as
 * PresenceLayer:
 *
 *   - Find the univer-host + main grid canvas.
 *   - Translate the chart's cell-coordinate pos into a host-local
 *     CSS box: `getCellRect(corner)` gives canvas-local pre-scroll
 *     coords; subtract `scrollX/scrollY` (tracked via Univer's
 *     Scroll event), add the canvas-vs-host offset.
 *   - Only render charts whose sheet matches the active sheet
 *     (Excel hides charts on inactive sheets).
 *   - Rerun the math every animation frame so the chart sticks
 *     to its cell anchor through scroll + resize + zoom.
 *
 * Mounts into the univer-host via a portal so the overlay paints
 * on top of the canvas at the same z-stack as PresenceLayer.
 */
type RenderedChart = {
  id: string;
  rect: { left: number; top: number; width: number; height: number };
};

type CtxMenuState = { id: string; x: number; y: number } | null;

export function ChartLayer() {
  const { api: sdkApi, charts, selectedId, select, remove } = useCharts();
  // FUniver facade — reached through the SDK handle (the app used a
  // `useUniverAPI()` hook that doesn't exist in the packaged SDK).
  const api = sdkApi.univer;
  const [rendered, setRendered] = useState<RenderedChart[]>([]);
  const [ctxMenu, setCtxMenu] = useState<CtxMenuState>(null);
  const [canvasOffset, setCanvasOffset] = useState({ x: 0, y: 0 });
  // rAF callback closes over `rendered` from when the effect ran. The effect
  // re-runs only on [api, charts] changes, so a sheet-switch (which leaves
  // both inputs untouched) would compare new computed rects against the
  // stale closure copy — `rectsEqual` could return true while React state
  // still holds the previous overlays. Mirror state into a ref so the
  // diff always sees the live value.
  const renderedRef = useRef<RenderedChart[]>(rendered);
  renderedRef.current = rendered;
  const hostRef = useRef<HTMLElement | null>(null);
  const scrollRef = useRef({ x: 0, y: 0 });

  useEffect(() => {
    hostRef.current = getUniverHost();
  }, []);

  // The chart's screen position is the cell's canvas-local rect minus
  // the viewport scroll offset. `Event.Scroll` would be the natural
  // hook, but it doesn't fire on programmatic scrolls (scrollToCell)
  // and the registration timing is fragile against lifecycle stages.
  // `getScrollState()` is the authoritative read — poll it every
  // animation frame and diff. Cheap (a few number reads) and correct.

  useEffect(() => {
    if (!api) return;
    let raf = 0;

    const recompute = () => {
      const host = hostRef.current ?? getUniverHost();
      if (!host) {
        if (renderedRef.current.length) setRendered([]);
        return;
      }
      const canvas = getUniverMainCanvas(host);
      if (!canvas) {
        if (renderedRef.current.length) setRendered([]);
        return;
      }
      const hostRect = host.getBoundingClientRect();
      const canvasRect = canvas.getBoundingClientRect();
      const dx = canvasRect.left - hostRect.left;
      const dy = canvasRect.top - hostRect.top;
      if (dx !== canvasOffset.x || dy !== canvasOffset.y) {
        setCanvasOffset({ x: dx, y: dy });
      }

      const wb = api.getActiveWorkbook();
      if (!wb) {
        if (renderedRef.current.length) setRendered([]);
        return;
      }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const activeSheet = wb.getActiveSheet() as any;
      const activeSheetId = activeSheet?.getSheetId?.();

      // Read viewport scroll directly from the worksheet — each frame
      // is cheap (a handful of property reads) and dodges Univer's
      // flaky `Event.Scroll` dispatch (doesn't fire on `scrollToCell`,
      // and the listener registration races the render lifecycle).
      // Cell rects come back canvas-local + pre-scroll; we convert by
      // asking for the rect of the cell currently at viewport top-left
      // — that rect's `top/left` IS the scroll offset to subtract.
      // CRITICAL: `getScrollState()` itself can throw a redi
      // QuantityCheckError during the brief window between Univer mount
      // and full render-unit DI graph setup. The whole block —
      // including the getScrollState() call — must be inside try/catch.
      // Otherwise the rAF tick throws every frame until init completes.
      let sx = 0;
      let sy = 0;
      try {
        const scrollState = activeSheet?.getScrollState?.() as
          | {
              sheetViewStartRow?: number;
              sheetViewStartColumn?: number;
              offsetX?: number;
              offsetY?: number;
            }
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
        /* skeleton / scroll service not ready — leave scroll at 0 this frame */
      }
      // Stash the latest scroll for downstream consumers (ChartOverlay).
      // We no longer need a tick counter to gate recomputes — they fire
      // every frame now.
      scrollRef.current = { x: sx, y: sy };

      // Header gutter — see univer-dom.getHeaderGutter. Without this,
      // charts sit ~40 px above and left of their cell anchor.
      const gutter = getHeaderGutter(api);

      // Zoom — getCellRect returns logical content coords; Univer
      // applies zoom as a scene transform when drawing the canvas.
      // Multiply (content - scroll) by zoom so the overlay tracks
      // the canvas at any zoom level. See PresenceLayer for the
      // same pattern + rationale.
      let zoom = 1;
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ws = activeSheet as any;
        const z =
          (ws?._worksheet?.getZoomRatio?.() as number | undefined) ??
          (ws?.getZoomRatio?.() as number | undefined);
        if (typeof z === 'number' && z > 0) zoom = z;
      } catch {
        /* zoom unreadable — leave at 1 */
      }

      const out: RenderedChart[] = [];
      for (const c of charts) {
        if (c.sheetId !== activeSheetId) continue;
        try {
          const tl = activeSheet.getRange(c.pos.startRow, c.pos.startColumn).getCellRect();
          const br = activeSheet.getRange(c.pos.endRow, c.pos.endColumn).getCellRect();
          if (!tl || !br) continue;
          const left = (Math.min(tl.left, br.left) - sx) * zoom + dx + gutter.rowHeaderWidth;
          const top = (Math.min(tl.top, br.top) - sy) * zoom + dy + gutter.columnHeaderHeight;
          const right = (Math.max(tl.right, br.right) - sx) * zoom + dx + gutter.rowHeaderWidth;
          const bottom =
            (Math.max(tl.bottom, br.bottom) - sy) * zoom + dy + gutter.columnHeaderHeight;
          // Clip — charts mostly off-canvas don't render (saves
          // ECharts a redraw); partially off is fine because the
          // overlay overflow:hidden on the layer below clips it.
          if (right < dx || bottom < dy) continue;
          if (left > dx + canvasRect.width || top > dy + canvasRect.height) continue;
          out.push({
            id: c.id,
            rect: { left, top, width: right - left, height: bottom - top },
          });
        } catch {
          /* getCellRect can throw mid-resize — drop this frame */
        }
      }
      if (rectsEqual(out, renderedRef.current)) return;
      setRendered(out);
    };

    const tick = () => {
      // Recompute every animation frame. The cost is dominated by a
      // handful of `getCellRect` calls per chart and one `getHeaderGutter`
      // lookup — easily under a millisecond for typical workbooks.
      // Throttling to every 4 frames was the source of charts visually
      // pinning to the viewport for ~50 ms at the start of any scroll
      // (same root cause as the remote-cursor scroll-pin bug —
      // see docs/COLLAB-FIXES.md #14).
      recompute();
      raf = requestAnimationFrame(tick);
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
    // charts + api are the inputs; rendered is the output.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api, charts]);

  // Delete key removes the selected chart. Excel uses Delete/Backspace
  // both — match that. We ignore the press if the focus is in a text
  // input (formula bar, cell editor, dialog inputs) so the user can
  // still delete characters there. Also Esc clears selection.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!selectedId) return;
      const t = e.target as HTMLElement | null;
      const tag = t?.tagName;
      const inText = tag === 'INPUT' || tag === 'TEXTAREA' || (t?.isContentEditable ?? false);
      if (e.key === 'Escape') {
        select(null);
        return;
      }
      if (!inText && (e.key === 'Delete' || e.key === 'Backspace')) {
        e.preventDefault();
        remove(selectedId);
      }
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [selectedId, remove, select]);

  // Listen for the right-click event ChartOverlay dispatches. We can't
  // mount the context menu inside ChartOverlay because the menu needs
  // to escape the overlay's clipping + sit at viewport coords.
  useEffect(() => {
    const onCtx = (e: Event) => {
      const ce = e as CustomEvent<{ id: string; x: number; y: number }>;
      setCtxMenu(ce.detail);
    };
    document.addEventListener('casual-chart-contextmenu', onCtx);
    return () => document.removeEventListener('casual-chart-contextmenu', onCtx);
  }, []);

  // Click anywhere outside any chart deselects. Capture-phase so the
  // grid canvas doesn't consume the event first. Skip if the click
  // originated inside a chart overlay or the context menu.
  useEffect(() => {
    const onClick = (e: MouseEvent) => {
      const t = e.target as HTMLElement | null;
      if (!t) return;
      if (t.closest('.chart-overlay')) return;
      if (t.closest('.chart-context-menu')) return;
      if (selectedId) select(null);
    };
    document.addEventListener('mousedown', onClick, true);
    return () => document.removeEventListener('mousedown', onClick, true);
  }, [selectedId, select]);

  const host = hostRef.current ?? getUniverHost();
  if (rendered.length === 0 && !ctxMenu) return null;
  if (!host) return null;

  return createPortal(
    <div
      className="chart-layer"
      data-testid="chart-layer"
      // z-index matches presence-layer so chart overlays sit above
      // Univer's main grid canvas (which has its own stacking context).
      // pointerEvents: none keeps clicks on empty space passing through
      // to the canvas; .chart-overlay re-enables pointer events.
      style={{
        position: 'absolute',
        inset: 0,
        overflow: 'hidden',
        pointerEvents: 'none',
        zIndex: 60,
      }}
    >
      {rendered.map((r) => {
        const model = charts.find((c) => c.id === r.id);
        if (!model) return null;
        // ChartOverlay positions itself absolutely (left/top/w/h from
        // `r.rect`). Render it directly — the previous wrapper used
        // `inset: 0` which covered the ENTIRE host with pointer-events
        // and intercepted clicks meant for empty grid cells.
        return (
          <ChartOverlay
            key={r.id}
            model={model}
            rect={r.rect}
            canvasOffset={canvasOffset}
            scroll={scrollRef.current}
          />
        );
      })}
      {ctxMenu && (
        <ChartContextMenu
          chartId={ctxMenu.id}
          x={ctxMenu.x}
          y={ctxMenu.y}
          onClose={() => setCtxMenu(null)}
        />
      )}
    </div>,
    host,
  );
}

function rectsEqual(a: RenderedChart[], b: RenderedChart[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) {
    const x = a[i];
    const y = b[i];
    if (x.id !== y.id) return false;
    if (
      x.rect.left !== y.rect.left ||
      x.rect.top !== y.rect.top ||
      x.rect.width !== y.rect.width ||
      x.rect.height !== y.rect.height
    ) {
      return false;
    }
  }
  return true;
}
