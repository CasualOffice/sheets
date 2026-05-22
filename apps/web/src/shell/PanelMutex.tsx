import { useEffect, useRef } from 'react';
import { ISidebarService } from '@univerjs/ui';
import { useUniverAPI } from '../use-univer';
import { useUI } from '../use-ui';

/**
 * Keeps "only one right-side panel is open" true across the two panel
 * systems we have:
 *
 *   - React side panels (Tables / Charts / Outline / History) — our
 *     own state, controlled via ui-context.
 *   - Univer's built-in sidebar (Comments today; data-validation popup
 *     and others later) — managed by `ISidebarService`.
 *
 * Without this, opening Comments leaves a previously-open React panel
 * visible and the user gets two stacked sidebars fighting for the same
 * width. The fix is two-way:
 *
 *   1. When ANY React panel becomes visible → call `sidebarService.close()`
 *      so any Univer sidebar already open dismisses.
 *   2. When `sidebarOptions$` emits with `visible: true` → call
 *      `ui.closeAllReactPanels()`.
 *
 * IMPORTANT: this listener must resolve the sidebar service using the
 * real `ISidebarService` identifier (imported from @univerjs/ui), NOT
 * the string `'ui.sidebar.service'`. Univer's redi injector keys on
 * the identifier object's internal Symbol; the string id is only for
 * debugging — `injector.get('ui.sidebar.service')` returns undefined.
 *
 * The `lastVisibleRef` guards against the self-echo: closing the
 * sidebar emits with `visible: false` (and the previous panel's
 * `children`), so we only act on the rising edge (false → true).
 */
export function PanelMutex() {
  const api = useUniverAPI();
  const ui = useUI();

  // Step 1 — close Univer's sidebar whenever a React panel becomes visible.
  const anyReactPanelOpen =
    ui.tablesPanelVisible ||
    ui.chartsPanelVisible ||
    ui.outlinePanelVisible ||
    ui.historyPanelVisible;
  useEffect(() => {
    if (!api || !anyReactPanelOpen) return;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const injector = (api as any)._injector as
        | { get: (token: unknown) => unknown }
        | undefined;
      if (!injector) return;
      const sidebarService = injector.get(ISidebarService) as
        | { close: () => void; visible: boolean }
        | undefined;
      if (sidebarService?.visible) sidebarService.close();
    } catch (err) {
      console.debug('[panel-mutex] could not reach sidebar service', err);
    }
  }, [api, anyReactPanelOpen]);

  // Step 2 — subscribe to Univer's sidebar opens; close React panels
  // on the rising edge. The Subject re-emits on every open/close, so
  // we filter to "false → true" transitions to avoid the self-loop.
  const lastVisibleRef = useRef(false);
  useEffect(() => {
    if (!api) return;
    let sub: { unsubscribe: () => void } | undefined;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const injector = (api as any)._injector as
        | { get: (token: unknown) => unknown }
        | undefined;
      if (!injector) return;
      const sidebarService = injector.get(ISidebarService) as
        | {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            sidebarOptions$: { subscribe: (fn: (o: any) => void) => { unsubscribe: () => void } };
            readonly visible: boolean;
          }
        | undefined;
      if (!sidebarService) return;
      // Seed `lastVisibleRef` so a sidebar that's already open at
      // mount time doesn't trigger an immediate React-panel sweep.
      lastVisibleRef.current = sidebarService.visible;
      sub = sidebarService.sidebarOptions$.subscribe((opts) => {
        const nowVisible = Boolean(opts && opts.visible);
        if (nowVisible && !lastVisibleRef.current) {
          // Univer just opened the sidebar. Close every React panel
          // so the right edge has one occupant only.
          ui.closeAllReactPanels();
        }
        lastVisibleRef.current = nowVisible;
      });
    } catch (err) {
      console.debug('[panel-mutex] could not subscribe to sidebar', err);
    }
    return () => sub?.unsubscribe();
  }, [api, ui]);

  return null;
}
