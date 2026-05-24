import { useEffect } from 'react';

/**
 * Touch-pan adapter for the Univer canvas.
 *
 * Univer 0.24 has no native touch-pan: its viewport scrolling is wired
 * exclusively to `wheel` events, so on mobile a swipe does nothing. The
 * canvas is `touch-action: none` (set by Univer itself), so the browser
 * also won't pan the page for us.
 *
 * This hook listens to touchmove on the document in capture phase. When
 * a single-finger touch on a Univer render-canvas drags more than the
 * tap-vs-drag threshold, we:
 *
 *   1. Stop the touch from reaching Univer's pointer-down/move handlers
 *      (those would interpret the drag as a cell-selection extend, so
 *      the user would scroll AND extend a selection at the same time).
 *   2. Dispatch a synthetic `WheelEvent` at the canvas with the delta,
 *      which Univer's existing wheel handler translates to a viewport
 *      scroll — the same code path desktop uses.
 *
 * Short stationary taps fall through unchanged, so tap-to-select still
 * works exactly like desktop click-to-select. Two-finger gestures are
 * left alone (Univer's input manager handles them; users get native
 * browser pinch-zoom for scaling when applicable).
 *
 * If/when Univer adds first-class touch-pan to its viewport, drop this
 * hook entirely.
 */

const CANVAS_SELECTOR = '[data-u-comp="render-canvas"]';
const PAN_THRESHOLD_PX = 6;

export function useTouchPan(): void {
  useEffect(() => {
    let canvas: HTMLCanvasElement | null = null;
    let startX = 0;
    let startY = 0;
    let lastX = 0;
    let lastY = 0;
    let panning = false;

    const onTouchStart = (e: TouchEvent) => {
      if (e.touches.length !== 1) {
        // Two-finger gestures — leave them for Univer / browser pinch.
        canvas = null;
        panning = false;
        return;
      }
      const t = e.touches[0];
      // closest() walks up from the touch target — works whether the
      // touch lands on the canvas itself or an overlaid Univer element.
      const target = (e.target as HTMLElement | null)?.closest(CANVAS_SELECTOR) as
        | HTMLCanvasElement
        | null;
      if (!target) {
        canvas = null;
        return;
      }
      canvas = target;
      startX = lastX = t.clientX;
      startY = lastY = t.clientY;
      panning = false;
    };

    const onTouchMove = (e: TouchEvent) => {
      if (!canvas || e.touches.length !== 1) return;
      const t = e.touches[0];

      if (!panning) {
        // Wait until movement crosses the tap-vs-drag threshold before
        // committing to pan mode. Below the threshold, a stationary or
        // near-stationary touch counts as a tap and falls through to
        // Univer's pointer handler unchanged.
        const totalDx = t.clientX - startX;
        const totalDy = t.clientY - startY;
        if (Math.abs(totalDx) < PAN_THRESHOLD_PX && Math.abs(totalDy) < PAN_THRESHOLD_PX) {
          return;
        }
        panning = true;
      }

      // Stop the touch from propagating to Univer's pointer-move
      // handler. Otherwise Univer interprets the drag as a cell-
      // selection extend AND we scroll — two motions at once.
      e.stopImmediatePropagation();
      e.preventDefault();

      const dx = lastX - t.clientX;
      const dy = lastY - t.clientY;
      lastX = t.clientX;
      lastY = t.clientY;

      // Synthesise a wheel event at the canvas. Univer's _pointerWheelEvent
      // listens on the canvas; this is exactly what a trackpad two-finger
      // scroll would dispatch.
      const wheel = new WheelEvent('wheel', {
        deltaX: dx,
        deltaY: dy,
        deltaMode: 0, // pixel units
        bubbles: true,
        cancelable: true,
        clientX: t.clientX,
        clientY: t.clientY,
      });
      canvas.dispatchEvent(wheel);
    };

    const onTouchEnd = () => {
      canvas = null;
      panning = false;
    };

    // Capture phase so we get the touch before Univer's pointer handlers.
    // `passive: false` is required to call preventDefault on touchmove.
    document.addEventListener('touchstart', onTouchStart, { capture: true, passive: true });
    document.addEventListener('touchmove', onTouchMove, { capture: true, passive: false });
    document.addEventListener('touchend', onTouchEnd, { capture: true, passive: true });
    document.addEventListener('touchcancel', onTouchEnd, { capture: true, passive: true });

    return () => {
      document.removeEventListener('touchstart', onTouchStart, { capture: true } as EventListenerOptions);
      document.removeEventListener('touchmove', onTouchMove, { capture: true } as EventListenerOptions);
      document.removeEventListener('touchend', onTouchEnd, { capture: true } as EventListenerOptions);
      document.removeEventListener('touchcancel', onTouchEnd, { capture: true } as EventListenerOptions);
    };
  }, []);
}
