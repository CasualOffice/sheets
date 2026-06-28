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

import { useEffect } from 'react';

/**
 * Touch-pan adapter for the Univer canvas.
 *
 * Univer 0.24 has no native touch-pan: its viewport scrolling is wired
 * exclusively to `wheel` events, so on mobile a swipe does nothing. The
 * canvas is `touch-action: none` (set by Univer itself), so the browser
 * also won't pan the page for us.
 *
 * We listen on POINTER events (pointertype = 'touch') in capture phase,
 * not touch events. Univer's drag-to-select runs off `pointermove` on
 * the canvas — touch events and pointer events are SEPARATE streams, so
 * blocking touchmove doesn't stop pointermove from extending the
 * selection. (First version of this hook used touch events; the
 * symptom was "horizontal swipe scrolls + selects, vertical swipe
 * selects only" because the selection extension visually dominated.)
 *
 * Flow:
 *   1. pointerdown (touch) — note the canvas + start coords. Don't
 *      block — Univer needs to register the initial cell for tap-
 *      to-select. If the user lifts without moving past the
 *      threshold, that tap-select stands.
 *   2. pointermove (touch) — once movement crosses the tap-vs-drag
 *      threshold, stopImmediatePropagation so Univer's pointermove
 *      handler never fires (no rogue selection-extend), then
 *      dispatch a synthetic WheelEvent at the canvas with the delta.
 *      Univer's wheel handler scrolls the viewport via the same
 *      code path desktop uses.
 *   3. pointerup / pointercancel — clear state.
 *
 * Mouse and pen pointers fall through unchanged so desktop behaviour
 * is identical.
 *
 * If/when Univer ships first-class touch-pan upstream, delete this.
 */

const CANVAS_SELECTOR = '[data-u-comp="render-canvas"]';
const PAN_THRESHOLD_PX = 6;

export function useTouchPan(): void {
  useEffect(() => {
    let canvas: HTMLCanvasElement | null = null;
    let activePointerId: number | null = null;
    let startX = 0;
    let startY = 0;
    let lastX = 0;
    let lastY = 0;
    let panning = false;

    const onPointerDown = (e: PointerEvent) => {
      if (e.pointerType !== 'touch') return;
      // Only track one finger at a time. Two-finger gestures are
      // left to Univer / the browser (pinch-zoom etc.).
      if (activePointerId !== null) return;
      const target = (e.target as HTMLElement | null)?.closest(CANVAS_SELECTOR) as
        | HTMLCanvasElement
        | null;
      if (!target) return;
      canvas = target;
      activePointerId = e.pointerId;
      startX = lastX = e.clientX;
      startY = lastY = e.clientY;
      panning = false;
    };

    const onPointerMove = (e: PointerEvent) => {
      if (
        !canvas ||
        e.pointerType !== 'touch' ||
        e.pointerId !== activePointerId
      )
        return;

      if (!panning) {
        // Wait until movement crosses the tap-vs-drag threshold
        // before committing to pan mode. Below the threshold, fall
        // through unchanged so a stationary tap reaches Univer's
        // pointer handler and the cell selects normally.
        const totalDx = e.clientX - startX;
        const totalDy = e.clientY - startY;
        if (
          Math.abs(totalDx) < PAN_THRESHOLD_PX &&
          Math.abs(totalDy) < PAN_THRESHOLD_PX
        ) {
          return;
        }
        panning = true;
      }

      // CRITICAL: stop the pointer move from reaching Univer's own
      // pointermove handler. Without this, Univer reads the move as
      // a cell-selection extend and the user gets a phantom grid
      // selection on every vertical swipe.
      e.stopImmediatePropagation();
      e.preventDefault();

      const dx = lastX - e.clientX;
      const dy = lastY - e.clientY;
      lastX = e.clientX;
      lastY = e.clientY;

      // Synthesise a wheel at the canvas. Same shape as a trackpad
      // two-finger scroll — Univer's _pointerWheelEvent picks it up
      // and translates to a viewport scroll.
      const wheel = new WheelEvent('wheel', {
        deltaX: dx,
        deltaY: dy,
        deltaMode: 0,
        bubbles: true,
        cancelable: true,
        clientX: e.clientX,
        clientY: e.clientY,
      });
      canvas.dispatchEvent(wheel);
    };

    const onPointerEnd = (e: PointerEvent) => {
      if (e.pointerId !== activePointerId) return;
      // If a pan was in progress, swallow the final pointerup so
      // Univer doesn't interpret the gesture as a click.
      if (panning) {
        e.stopImmediatePropagation();
      }
      canvas = null;
      activePointerId = null;
      panning = false;
    };

    document.addEventListener('pointerdown', onPointerDown, { capture: true });
    document.addEventListener('pointermove', onPointerMove, { capture: true, passive: false });
    document.addEventListener('pointerup', onPointerEnd, { capture: true });
    document.addEventListener('pointercancel', onPointerEnd, { capture: true });

    return () => {
      document.removeEventListener('pointerdown', onPointerDown, { capture: true } as EventListenerOptions);
      document.removeEventListener('pointermove', onPointerMove, { capture: true } as EventListenerOptions);
      document.removeEventListener('pointerup', onPointerEnd, { capture: true } as EventListenerOptions);
      document.removeEventListener('pointercancel', onPointerEnd, { capture: true } as EventListenerOptions);
    };
  }, []);
}
