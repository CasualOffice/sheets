/**
 * Helpers for reaching into Univer's mounted DOM. Internal Univer DOM
 * identifiers are not part of the public API and have shifted between
 * minor versions — centralise the selectors here so a Univer upgrade
 * needs one fix instead of touching every overlay that anchors to the
 * canvas (PresenceLayer, ChartLayer, future drawing/comment overlays).
 *
 * Every helper falls back to a structural query and warns once if it
 * had to use the fallback, so a silent overlay-disappears regression
 * surfaces in the console instead of looking like a sync bug.
 */

import type { FUniver } from '@univerjs/core/facade';

const PRIMARY_CANVAS_SELECTOR = 'canvas[id^="univer-sheet-main-canvas_"]';
/** Univer mounts a hidden formula-editor canvas BEFORE the main grid
 *  canvas. That one has no id, sits with width/height = 0, and is
 *  useless to anchor overlays to. The fallback path must explicitly
 *  skip it — picking it caused PresenceLayer / ChartLayer to render
 *  cursors at the wrong position whenever the primary selector raced
 *  the main-canvas mount. */
let fallbackWarned = false;

export function getUniverHost(): HTMLElement | null {
  return document.querySelector('[data-testid="univer-host"]') as HTMLElement | null;
}

export function getUniverMainCanvas(host: HTMLElement): HTMLCanvasElement | null {
  const primary = host.querySelector(PRIMARY_CANVAS_SELECTOR) as HTMLCanvasElement | null;
  if (primary) return primary;
  // Walk all canvases and pick the first one that LOOKS like the main
  // grid: has a non-empty id AND a non-zero rendered size. Editor
  // canvases have empty ids and 0x0 boxes until focused. We refuse to
  // return any of those — better to render no overlay than to anchor
  // it to a hidden offscreen canvas.
  const all = Array.from(host.querySelectorAll('canvas')) as HTMLCanvasElement[];
  for (const c of all) {
    if (!c.id) continue;
    const r = c.getBoundingClientRect();
    if (r.width < 50 || r.height < 50) continue;
    if (!fallbackWarned) {
      fallbackWarned = true;
      console.warn(
        '[univer-dom] primary canvas selector "%s" matched nothing — using fallback id "%s". Update PRIMARY_CANVAS_SELECTOR after a Univer upgrade.',
        PRIMARY_CANVAS_SELECTOR,
        c.id,
      );
    }
    return c;
  }
  return null;
}

/**
 * Pixel offset between the canvas top-left and the cell-content area's
 * top-left — i.e. the row-header gutter width and the column-header
 * gutter height. Every overlay that anchors to `getCellRect()`
 * coordinates must add these to land on the actual cells instead of
 * ~40 px up-and-left of them.
 *
 * We use Univer 0.22.x's documented defaults (row header = 46 px,
 * column header = 20 px) rather than reaching into the render
 * skeleton at runtime — the dynamic lookup via
 * `RenderUnit.with(SheetSkeletonManagerService)` instantiates the
 * service through redi, which transitively requires
 * `SheetScrollManagerService` that may not yet be registered on the
 * unit's injector when our rAF tick fires (race with Univer's
 * render-unit init). The crash surfaces as
 * "QuantityCheckError: Expect 1 dependency item(s) for id
 * 'SheetScrollManagerService' but get 0".
 *
 * Hardcoding the defaults trades ~40 px accuracy for any user who
 * customises header sizes (rare, undocumented in our app) for
 * zero-risk init. If you change Univer's default theme, update
 * these constants and the cursor/chart layers stay aligned.
 */
export type HeaderGutter = { rowHeaderWidth: number; columnHeaderHeight: number };

const DEFAULT_HEADER_GUTTER: HeaderGutter = {
  rowHeaderWidth: 46,
  columnHeaderHeight: 20,
};

export function getHeaderGutter(_api: FUniver | null): HeaderGutter {
  return DEFAULT_HEADER_GUTTER;
}
