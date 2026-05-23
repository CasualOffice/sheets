import { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { useUniverAPI } from '../use-univer';
import { getHeaderGutter, getUniverHost, getUniverMainCanvas } from '../univer-dom';
import { usePresence } from './presence-context';

/**
 * Renders remote peer selections as absolutely-positioned rectangles over
 * the grid. Strategy:
 *
 *   - Read the host (`[data-testid="univer-host"]`) bounding rect to know
 *     the overlay's reference frame.
 *   - For each peer's selection, call `range.getCellRect()` on its corner
 *     cells to derive *content-space* coords (not viewport — those are
 *     pre-scroll positions inside the canvas's drawing space).
 *   - Subtract the current scroll offset (tracked from `api.Event.Scroll`)
 *     and add the canvas-vs-host offset so positions land in the portal's
 *     coordinate frame.
 *   - Rerun on a ticker (rAF every 4 frames ≈ 67ms) so selection updates
 *     and small resize wobbles still settle quickly.
 *
 * The overlay div has `pointer-events: none` so it never blocks cell
 * clicks. CSS handles the color + label styling.
 */

type Rect = {
  clientId: number;
  name: string;
  color: string;
  left: number;
  top: number;
  width: number;
  height: number;
  /** Live in-progress text for this peer's cell, if any. Renders as a
   *  ghost overlay inside the cursor rect. */
  liveText?: string;
  /** Vertical offset (px) for the floating name label above the
   *  cursor. Default is -20 (one row above the cell). When two peers
   *  sit on adjacent cells their labels overlap horizontally —
   *  collision avoidance bumps later peers up by multiples of
   *  LABEL_SLOT_HEIGHT so they stack instead of stomping. */
  labelTop: number;
  /** Cell-anchor key (`unitId:sheetId:sr:sc:er:ec`). When this string
   *  differs from the previous frame's value, the peer's selection
   *  moved to a new range — we apply the `--moving` transition class
   *  so the rectangle eases between cells. When it matches, only the
   *  on-screen position is changing (e.g. user scrolling), so we
   *  paint instantly without a transition. */
  anchorKey: string;
  /** True for the non-primary ranges in a multi-range (Ctrl-click)
   *  selection. We omit the name label + live-edit ghost on these so
   *  the chrome doesn't multiply, and render them with a lighter fill
   *  so the primary still reads as "where the peer is right now". */
  isExtra?: boolean;
};

const LABEL_HEIGHT = 18;
const LABEL_SLOT_HEIGHT = 20;
const LABEL_BASE_TOP = -20;
/** Rough px-per-character for label width estimation. Lets us detect
 *  collisions without measuring DOM (the layer rebuilds every frame,
 *  so DOM measurement would thrash). */
const LABEL_CHAR_WIDTH = 6.5;
const LABEL_PAD = 12;

export function PresenceLayer() {
  const api = useUniverAPI();
  const { peers } = usePresence();
  const [rects, setRects] = useState<Rect[]>([]);
  // The rAF closure below captures `rects` only when the effect runs (deps:
  // [api, peers]). On scroll-driven recomputes neither dep changes, so the
  // stale closure copy makes `rectsEqual(next, rects)` always wrong against
  // the original mount snapshot — every frame causes a setState even when
  // positions are stable. ChartLayer fixed this with the same pattern.
  const rectsRef = useRef<Rect[]>(rects);
  rectsRef.current = rects;
  const hostRef = useRef<HTMLElement | null>(null);
  const layerRef = useRef<HTMLDivElement | null>(null);
  // Latest scroll offset from Univer's grid. `getCellRect` returns
  // content-space coords; subtract this to land in canvas-visible space.
  // Stashed for debugging — the rAF loop reads scroll inline each frame
  // and no longer needs a tick counter to gate recomputes.
  const scrollRef = useRef<{ x: number; y: number }>({ x: 0, y: 0 });

  // Resolve the Univer host once, then on each render rebuild rects.
  useEffect(() => {
    hostRef.current = getUniverHost();
  }, []);

  // Scroll offset is read inline in `recompute` via `getScrollState()` —
  // we used to subscribe to `Event.Scroll` here, but that event doesn't
  // fire on `scrollToCell` (programmatic scrolls leak through) and the
  // facade-side listener registration races the render lifecycle so
  // even wheel scrolls would silently no-op until the next remount.
  // Polling the worksheet directly per frame is cheap and authoritative.

  useEffect(() => {
    if (!api) return;
    // No peers → nothing to draw. Skip the rAF loop entirely instead of
    // spinning getCellRect / getScrollState work every frame for an empty
    // overlay. This is the common case: every single-user session has
    // zero peers. The effect re-runs when `peers` changes, so the loop
    // starts as soon as someone joins.
    if (peers.length === 0) {
      if (rectsRef.current.length) setRects([]);
      return;
    }
    let raf = 0;
    const tick = () => {
      // Recompute every animation frame. Univer scrolls the grid by
      // repainting the canvas — the DOM doesn't emit a scroll event,
      // so our only way to track scroll position is to poll. The work
      // per frame is ~O(peers): two `getCellRect` calls + a getBCR per
      // peer plus a `getHeaderGutter`, all O(1). For typical rooms
      // (< 10 active peers) this is well under a millisecond per frame.
      // Throttling was previously to every 4 frames, but that pinned
      // remote cursors to the viewport for the first ~50 ms of any
      // scroll, on top of the CSS transition that pinned them for
      // another ~80 ms — see docs/COLLAB-FIXES.md #14.
      recompute();
      raf = requestAnimationFrame(tick);
    };
    const recompute = () => {
      const host = hostRef.current ?? getUniverHost();
      if (!host) {
        if (rectsRef.current.length) setRects([]);
        return;
      }
      // The main grid canvas — its viewport-relative position is the
      // reference frame `getCellRect()` returns coords in. Without the
      // canvas offset we'd anchor cursors at (0,0) of the document.
      const canvas = getUniverMainCanvas(host);
      if (!canvas) {
        if (rectsRef.current.length) setRects([]);
        return;
      }
      const hostRect = host.getBoundingClientRect();
      const canvasRect = canvas.getBoundingClientRect();
      // Offset between the canvas's top-left and the host's top-left —
      // this is what we need to add to cell-local coords to land in
      // host-local coords (the portal's coordinate frame).
      const dx = canvasRect.left - hostRect.left;
      const dy = canvasRect.top - hostRect.top;
      // Header gutter — getCellRect returns coords in cell-content space,
      // which starts AT (rowHeaderWidth, columnHeaderHeight) inside the
      // canvas. Without adding these back, every cursor sits ~40 px up
      // and to the left of the cell it's labelling.
      const gutter = getHeaderGutter(api);

      const wb = api.getActiveWorkbook();
      if (!wb) {
        if (rectsRef.current.length) setRects([]);
        return;
      }
      const activeSheet = wb.getActiveSheet();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const activeSheetId = (activeSheet as any)?.getSheetId?.() ?? (activeSheet as any)?.getId?.() ?? '';

      // See header note — poll `getScrollState()` and derive the pixel
      // offset from the cell currently at viewport top-left.
      //
      // CRITICAL: `getScrollState()` itself can throw a redi
      // QuantityCheckError ("Expect 1 dependency item(s) for id
      // 'SheetScrollManagerService' but get 0") during the brief
      // window between Univer mount and full render-unit DI graph
      // setup. The whole block, INCLUDING the getScrollState call,
      // must be inside a try/catch — otherwise the rAF loop throws,
      // surfaces as an uncaught error in dev tools, and trips Vite's
      // overlay. In production the throw also halts the tick, freezing
      // remote cursor tracking.
      let sx = 0;
      let sy = 0;
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const scrollState = (activeSheet as any)?.getScrollState?.() as
          | { sheetViewStartRow?: number; sheetViewStartColumn?: number; offsetX?: number; offsetY?: number }
          | undefined;
        if (scrollState) {
          const r = scrollState.sheetViewStartRow ?? 0;
          const c = scrollState.sheetViewStartColumn ?? 0;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const topLeft = (activeSheet as any).getRange(r, c).getCellRect();
          if (topLeft) {
            sx = topLeft.left + (scrollState.offsetX ?? 0);
            sy = topLeft.top + (scrollState.offsetY ?? 0);
          }
        }
      } catch {
        /* skeleton / scroll service not ready — leave scroll at 0 this frame */
      }
      // Stash latest scroll so other consumers (devtools, debugging hooks)
      // can read it without redoing the math. We no longer use it to gate
      // recompute timing — recompute fires every frame now.
      scrollRef.current = { x: sx, y: sy };

      // Frozen-pane split — cells with row < freezeRow stay fixed at the
      // top (don't apply Y-scroll); cells with col < freezeCol stay
      // fixed at the left (don't apply X-scroll). Without this, peer
      // cursors in frozen rows/cols drift with the rest of the grid.
      // `startRow / startColumn` is the index of the first NON-frozen
      // row/col. xSplit / ySplit confirm intent (0 means no freeze).
      let freezeRow = 0;
      let freezeCol = 0;
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const f = (activeSheet as any)?.getFreeze?.() as
          | { startRow?: number; startColumn?: number; xSplit?: number; ySplit?: number }
          | undefined;
        if (f) {
          if ((f.ySplit ?? 0) > 0 && (f.startRow ?? -1) > 0) freezeRow = f.startRow!;
          if ((f.xSplit ?? 0) > 0 && (f.startColumn ?? -1) > 0) freezeCol = f.startColumn!;
        }
      } catch {
        /* freeze config unreadable — treat as no freeze */
      }

      // Zoom — `getCellRect` returns LOGICAL (unzoomed) content coords
      // because the underlying skeleton stores logical row heights and
      // column widths. Univer applies the zoom as a scene transform
      // when drawing the canvas, so the on-screen pixel position of
      // cell (r, c) is `tl.left * zoom + dx + headerGutter`. Read
      // zoomRatio from the worksheet's internal model — there's no
      // facade getter as of Univer 0.22.x. Defaults to 1 (no zoom).
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

      const next: Rect[] = [];
      const ws = wb.getActiveSheet();
      if (!ws) {
        if (rectsRef.current.length) setRects([]);
        return;
      }
      for (const peer of peers) {
        // Mismatched unit ids are normal — every browser gets a random
        // unitId on workbook creation. We render any peer whose sheet id
        // matches the local active sheet (sheet ids ARE deterministic).
        // Prefer the live-edit cell when present so the ghost overlay
        // sits exactly on the cell being typed into.
        let liveText: string | undefined;
        let ranges: Array<{ sr: number; er: number; sc: number; ec: number }> = [];
        if (peer.liveEdit && peer.liveEdit.s === activeSheetId) {
          ranges = [{
            sr: peer.liveEdit.row,
            er: peer.liveEdit.row,
            sc: peer.liveEdit.col,
            ec: peer.liveEdit.col,
          }];
          liveText = peer.liveEdit.text;
        } else if (peer.selection && peer.selection.s === activeSheetId) {
          // Multi-range (Ctrl-click) selections ride on `sel.rs`;
          // single-range stays on `sel.r` for backwards-compat. Render
          // every range — name label + ghost only on the primary
          // (index 0) so the chrome doesn't multiply.
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const selAny = peer.selection as any;
          if (Array.isArray(selAny.rs) && selAny.rs.length > 0) {
            ranges = selAny.rs;
          } else {
            ranges = [peer.selection.r];
          }
        } else {
          continue;
        }
        for (let idx = 0; idx < ranges.length; idx += 1) {
          const { sr, er, sc, ec } = ranges[idx];
          try {
            const tl = ws.getRange(sr, sc).getCellRect();
            const br2 = ws.getRange(er, ec).getCellRect();
            if (!tl || !br2) continue;
            // Frozen-pane handling on the range's START row/col — see
            // the comment block from the earlier single-range pass.
            const inFrozenRow = sr < freezeRow;
            const inFrozenCol = sc < freezeCol;
            const ySub = inFrozenRow ? 0 : sy;
            const xSub = inFrozenCol ? 0 : sx;
            const left = (Math.min(tl.left, br2.left) - xSub) * zoom + dx + gutter.rowHeaderWidth;
            const top = (Math.min(tl.top, br2.top) - ySub) * zoom + dy + gutter.columnHeaderHeight;
            const right = (Math.max(tl.right, br2.right) - xSub) * zoom + dx + gutter.rowHeaderWidth;
            const bottom = (Math.max(tl.bottom, br2.bottom) - ySub) * zoom + dy + gutter.columnHeaderHeight;
            // Clip per-range so off-screen extras don't ghost into the
            // headers.
            if (right < dx || bottom < dy) continue;
            if (left > dx + canvasRect.width || top > dy + canvasRect.height) continue;
            const isPrimary = idx === 0;
            next.push({
              clientId: peer.clientId,
              name: peer.name,
              color: peer.color,
              left,
              top,
              width: right - left,
              height: bottom - top,
              // Only the primary range carries the live-edit ghost +
              // name label. Extras render as bare highlighted rects.
              liveText: isPrimary ? liveText : undefined,
              labelTop: LABEL_BASE_TOP,
              isExtra: !isPrimary,
              anchorKey: `${wb.getId()}:${activeSheetId}:${sr}:${sc}:${er}:${ec}:${idx}`,
            });
          } catch {
            /* getCellRect can throw mid-resize — drop this frame for that range */
          }
        }
      }

      assignLabelSlots(next);

      // Cheap diff: only setState when the rect set actually changed,
      // so we don't churn React 15× per second.
      if (rectsEqual(next, rectsRef.current)) return;
      setRects(next);
    };

    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
    // peers + api are the inputs; rects are the output we're managing.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api, peers]);

  // Per-peer previous anchor key. Used to set `.presence-cursor--moving`
  // for one render after the peer's selection moves to a new cell, so
  // we get a smooth cell-to-cell ease without lerping every scroll
  // frame. Persists across renders via a ref so the comparison sees the
  // last committed value.
  const prevAnchorRef = useRef<Map<number, string>>(new Map());

  if (rects.length === 0) return null;
  const host = hostRef.current ?? getUniverHost();
  if (!host) return null;

  // Snapshot which cursors should animate this render. Done outside the
  // map() so we update the prev-anchor map exactly once per render and
  // don't trigger React-StrictMode double-invoke weirdness.
  const animating = new Set<number>();
  const nextAnchors = new Map<number, string>();
  for (const r of rects) {
    const prev = prevAnchorRef.current.get(r.clientId);
    if (prev !== undefined && prev !== r.anchorKey) animating.add(r.clientId);
    nextAnchors.set(r.clientId, r.anchorKey);
  }
  prevAnchorRef.current = nextAnchors;

  return createPortal(
    <div
      ref={layerRef}
      className="presence-layer"
      data-testid="presence-layer"
      aria-hidden="true"
    >
      {rects.map((r, idx) => {
        const classes = ['presence-cursor'];
        if (r.liveText !== undefined) classes.push('presence-cursor--editing');
        if (animating.has(r.clientId)) classes.push('presence-cursor--moving');
        if (r.isExtra) classes.push('presence-cursor--extra');
        return (
          <div
            key={`${r.clientId}-${idx}`}
            className={classes.join(' ')}
            data-testid="presence-cursor"
            data-live={r.liveText !== undefined ? '1' : '0'}
            data-extra={r.isExtra ? '1' : '0'}
            style={
              {
                left: `${r.left}px`,
                top: `${r.top}px`,
                width: `${r.width}px`,
                height: `${r.height}px`,
                ['--presence-color' as string]: r.color,
              } as React.CSSProperties
            }
          >
            {!r.isExtra && (
              <span
                className="presence-cursor__label"
                style={{ top: `${r.labelTop}px` }}
              >
                {r.name}
              </span>
            )}
            {r.liveText !== undefined && r.liveText.length > 0 && (
              <span className="presence-cursor__ghost" data-testid="presence-cursor-ghost">
                {r.liveText}
              </span>
            )}
          </div>
        );
      })}
    </div>,
    host,
  );
}

function rectsEqual(a: Rect[], b: Rect[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) {
    const x = a[i];
    const y = b[i];
    if (
      x.clientId !== y.clientId ||
      x.left !== y.left ||
      x.top !== y.top ||
      x.width !== y.width ||
      x.height !== y.height ||
      x.name !== y.name ||
      x.color !== y.color ||
      x.liveText !== y.liveText ||
      x.labelTop !== y.labelTop ||
      x.anchorKey !== y.anchorKey ||
      (x.isExtra ?? false) !== (y.isExtra ?? false)
    ) {
      return false;
    }
  }
  return true;
}

/**
 * Greedy label-collision pass. Walks the cursor rects in display order
 * and, for each label, finds the first vertical "slot" above the cursor
 * cell that doesn't overlap any earlier-placed label's bounding box.
 * Labels stack upward in 20 px increments — matches Google Sheets and
 * Office Online behavior when peers cluster.
 *
 * Mutates the array in place. Runs O(n²) which is fine — a typical
 * collab session has < 10 active peers in view at once.
 */
function assignLabelSlots(rects: Rect[]): void {
  const placed: Array<{ left: number; right: number; top: number; bottom: number }> = [];
  // Sort by cursor top so the topmost cursor takes slot 0 — feels
  // more natural visually than "first peer in the map wins".
  const order = rects
    .map((_, i) => i)
    .sort((a, b) => rects[a].top - rects[b].top);
  for (const idx of order) {
    const r = rects[idx];
    // Extras don't carry a name label, so they don't need a slot.
    if (r.isExtra) continue;
    const labelLeft = r.left - 2;
    const labelRight = labelLeft + Math.ceil(r.name.length * LABEL_CHAR_WIDTH) + LABEL_PAD;
    let slot = 0;
    // Try increasing slots until we don't intersect any placed label.
    while (slot < 20) {
      const labelTopAbs = r.top + LABEL_BASE_TOP - slot * LABEL_SLOT_HEIGHT;
      const labelBottomAbs = labelTopAbs + LABEL_HEIGHT;
      const conflict = placed.some(
        (p) =>
          labelLeft < p.right &&
          labelRight > p.left &&
          labelTopAbs < p.bottom &&
          labelBottomAbs > p.top,
      );
      if (!conflict) {
        rects[idx] = { ...r, labelTop: LABEL_BASE_TOP - slot * LABEL_SLOT_HEIGHT };
        placed.push({
          left: labelLeft,
          right: labelRight,
          top: labelTopAbs,
          bottom: labelBottomAbs,
        });
        break;
      }
      slot += 1;
    }
  }
}
