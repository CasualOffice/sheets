import { useCallback, useEffect, useRef, useState } from 'react';
import { useUniverAPI } from '../use-univer';
import { useCharts } from './charts-context';
import { init, type EChartsType } from './echarts-init';
import { buildEChartsOption } from './build-option';
import { rectToCellPos } from './hit-test';
import type { ChartModel } from './types';

/**
 * Single chart rendered on screen. Responsibilities:
 *
 *   - Owns one ECharts instance, mounted into a dedicated div.
 *   - Reads source data on mount + on every workbook value change,
 *     redraws via `setOption(option, true)`.
 *   - Selection: click selects (frame + 8 resize handles appear);
 *     the parent (ChartLayer) handles click-outside-to-deselect.
 *   - Drag-to-move: pointer-down on the body, drag, release → snap
 *     the model.pos top-left to the nearest cell, keep the size.
 *   - Drag-to-resize: pointer-down on any of the 8 handles, drag,
 *     release → snap both anchor corners to cells.
 *
 * Excel parity notes:
 *   - 4 corner handles (resize both axes) + 4 mid-edge handles
 *     (resize one axis).
 *   - Cursor changes to nwse/nesw/ns/ew depending on handle.
 *   - Move/resize during drag is free-positioned via CSS; the cell
 *     anchor only updates on pointer-up (so the model stays in cell
 *     coordinates and round-trips cleanly through xlsx + collab).
 */
type Props = {
  model: ChartModel;
  /** Host-local CSS box. Parent computes from the cell rect + the
   *  active sheet's scroll offset on every animation frame. */
  rect: { left: number; top: number; width: number; height: number };
  /** Canvas-vs-host offset + current scroll. Needed to convert the
   *  chart's screen pixels back to canvas-local pre-scroll coords
   *  when snapping to cells on drop. */
  canvasOffset: { x: number; y: number };
  scroll: { x: number; y: number };
};

type Handle = 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w' | 'nw';

const CORNER_HANDLES: Handle[] = ['nw', 'ne', 'se', 'sw'];
const EDGE_HANDLES: Handle[] = ['n', 'e', 's', 'w'];
const ALL_HANDLES: Handle[] = [...CORNER_HANDLES, ...EDGE_HANDLES];

const HANDLE_CURSORS: Record<Handle, string> = {
  n: 'ns-resize',
  s: 'ns-resize',
  e: 'ew-resize',
  w: 'ew-resize',
  ne: 'nesw-resize',
  sw: 'nesw-resize',
  nw: 'nwse-resize',
  se: 'nwse-resize',
};

const MIN_W = 80;
const MIN_H = 60;

export function ChartOverlay({ model, rect, canvasOffset, scroll }: Props) {
  const api = useUniverAPI();
  const { selectedId, select, update } = useCharts();
  const hostRef = useRef<HTMLDivElement>(null);
  const echartRef = useRef<EChartsType | null>(null);
  const isSelected = selectedId === model.id;

  // `drag` is the ephemeral pixel-space delta applied while the user
  // is dragging the body or a handle. It's a CSS-only transform until
  // pointer-up, at which point we compute the new cell anchor and
  // commit via `update()`. Null when not dragging.
  const [drag, setDrag] = useState<null | {
    mode: 'move' | Handle;
    startMouse: { x: number; y: number };
    startRect: { left: number; top: number; width: number; height: number };
    current: { left: number; top: number; width: number; height: number };
  }>(null);

  // Init ECharts once + dispose on unmount. ECharts mutates its own
  // DOM children so it must own its container.
  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const inst = init(host, undefined, { renderer: 'canvas' });
    echartRef.current = inst;
    return () => {
      inst.dispose();
      echartRef.current = null;
    };
  }, []);

  // Resize whenever the cell-rect-derived box (or the ephemeral drag
  // rect) changes. ECharts won't auto-resize on its own.
  useEffect(() => {
    echartRef.current?.resize();
  }, [rect.width, rect.height, drag?.current.width, drag?.current.height]);

  useEffect(() => {
    if (!api) return;
    const refresh = () => {
      const opt = buildEChartsOption(api, model);
      if (opt) echartRef.current?.setOption(opt, true);
    };
    refresh();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const evName = (api as any).Event?.SheetValueChanged;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    if (!evName || typeof (api as any).addEvent !== 'function') return;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const disp = (api as any).addEvent(evName, refresh) as { dispose?: () => void };
    return () => disp.dispose?.();
  }, [api, model]);

  // Pointer-up handler installed at the document level while a drag
  // is in flight. We can't put it on the overlay because the pointer
  // can escape during the move (and Excel doesn't lock pointers either).
  const commitDrag = useCallback(() => {
    if (!drag || !api) return;
    const wb = api.getActiveWorkbook();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const sheets = wb?.getSheets() as any[] | undefined;
    const ws = sheets?.find((s) => s.getSheetId?.() === model.sheetId);
    if (!ws) {
      setDrag(null);
      return;
    }
    const newPos = rectToCellPos(ws, drag.current, canvasOffset, scroll, {
      startRow: model.pos.startRow,
      startColumn: model.pos.startColumn,
      endRow: model.pos.endRow,
      endColumn: model.pos.endColumn,
    });
    if (newPos) update(model.id, { pos: newPos });
    setDrag(null);
  }, [api, canvasOffset, drag, model, scroll, update]);

  useEffect(() => {
    if (!drag) return;
    const onMove = (e: PointerEvent) => {
      const dx = e.clientX - drag.startMouse.x;
      const dy = e.clientY - drag.startMouse.y;
      setDrag((cur) => {
        if (!cur) return cur;
        const next = computeDragRect(cur.mode, cur.startRect, dx, dy);
        return { ...cur, current: next };
      });
    };
    const onUp = () => commitDrag();
    document.addEventListener('pointermove', onMove);
    document.addEventListener('pointerup', onUp);
    document.addEventListener('pointercancel', onUp);
    return () => {
      document.removeEventListener('pointermove', onMove);
      document.removeEventListener('pointerup', onUp);
      document.removeEventListener('pointercancel', onUp);
    };
  }, [drag, commitDrag]);

  const onBodyPointerDown = (e: React.PointerEvent) => {
    if (e.button !== 0) return; // primary button only
    e.stopPropagation();
    select(model.id);
    // A bare click selects but doesn't start a drag — the drag effect
    // only commits on pointer-up if `drag.current` differs from
    // `drag.startRect`. So we always set drag and trust commit logic.
    setDrag({
      mode: 'move',
      startMouse: { x: e.clientX, y: e.clientY },
      startRect: { ...rect },
      current: { ...rect },
    });
  };

  const onHandlePointerDown = (handle: Handle) => (e: React.PointerEvent) => {
    if (e.button !== 0) return;
    e.stopPropagation();
    select(model.id);
    setDrag({
      mode: handle,
      startMouse: { x: e.clientX, y: e.clientY },
      startRect: { ...rect },
      current: { ...rect },
    });
  };

  const liveRect = drag?.current ?? rect;

  return (
    <div
      ref={hostRef}
      className={`chart-overlay${isSelected ? ' chart-overlay--selected' : ''}${drag ? ' chart-overlay--dragging' : ''}`}
      data-testid="chart-overlay"
      data-chart-id={model.id}
      data-selected={isSelected ? 'true' : undefined}
      style={{
        position: 'absolute',
        left: `${liveRect.left}px`,
        top: `${liveRect.top}px`,
        width: `${liveRect.width}px`,
        height: `${liveRect.height}px`,
        cursor: drag?.mode === 'move' ? 'grabbing' : isSelected ? 'grab' : 'pointer',
      }}
      onPointerDown={onBodyPointerDown}
      onContextMenu={(e) => {
        e.preventDefault();
        select(model.id);
        // ChartLayer listens for this event and pops a context menu —
        // we just need to ensure the chart is selected first.
        const ce = new CustomEvent('casual-chart-contextmenu', {
          detail: { id: model.id, x: e.clientX, y: e.clientY },
        });
        document.dispatchEvent(ce);
      }}
    >
      {isSelected &&
        ALL_HANDLES.map((h) => (
          <div
            key={h}
            data-testid={`chart-handle-${h}`}
            className={`chart-overlay__handle chart-overlay__handle--${h}`}
            style={{ cursor: HANDLE_CURSORS[h] }}
            onPointerDown={onHandlePointerDown(h)}
          />
        ))}
    </div>
  );
}

/**
 * Apply a drag delta to the starting rect based on which handle (or
 * the body) is being dragged. Mode `'move'` slides the rect; the
 * handle modes resize from the matching edge / corner. We clamp
 * each dimension to a minimum so the chart can't collapse onto a
 * point (impossible to grab again).
 */
function computeDragRect(
  mode: 'move' | Handle,
  start: { left: number; top: number; width: number; height: number },
  dx: number,
  dy: number,
): { left: number; top: number; width: number; height: number } {
  if (mode === 'move') {
    return { left: start.left + dx, top: start.top + dy, width: start.width, height: start.height };
  }
  let { left, top, width, height } = start;
  const wantsLeft = mode === 'nw' || mode === 'w' || mode === 'sw';
  const wantsRight = mode === 'ne' || mode === 'e' || mode === 'se';
  const wantsTop = mode === 'nw' || mode === 'n' || mode === 'ne';
  const wantsBottom = mode === 'sw' || mode === 's' || mode === 'se';
  if (wantsLeft) {
    const newWidth = Math.max(MIN_W, start.width - dx);
    left = start.left + (start.width - newWidth);
    width = newWidth;
  } else if (wantsRight) {
    width = Math.max(MIN_W, start.width + dx);
  }
  if (wantsTop) {
    const newHeight = Math.max(MIN_H, start.height - dy);
    top = start.top + (start.height - newHeight);
    height = newHeight;
  } else if (wantsBottom) {
    height = Math.max(MIN_H, start.height + dy);
  }
  return { left, top, width, height };
}
