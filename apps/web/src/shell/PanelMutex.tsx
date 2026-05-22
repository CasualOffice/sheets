import { useEffect, useRef } from 'react';
import { useUniverAPI } from '../use-univer';
import { useUI } from '../use-ui';

/**
 * Keeps "only one right-side panel is open" true ACROSS the two panel
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
 *   - When ANY React panel becomes visible → call `sidebarService.close()`
 *     so any Univer sidebar already open dismisses.
 *   - When `sidebarService.sidebarOptions$` emits with a non-empty body
 *     (Univer just opened a sidebar) → call `ui.closeAllReactPanels()`.
 *
 * The `lastSidebarVisible` ref guards against the close-triggered self
 * echo (we don't want our own close to fire the React-panel sweep).
 */
export function PanelMutex() {
  const api = useUniverAPI();
  const ui = useUI();

  // Close Univer's sidebar whenever a React panel becomes visible.
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
      // ISidebarService identifier; resolve by its string id so we don't
      // have to import @univerjs/ui here (kept light to keep this module
      // out of the critical render path).
      const sidebarService = injector.get('ui.sidebar.service') as
        | { close: () => void; visible: boolean }
        | undefined;
      if (sidebarService?.visible) sidebarService.close();
    } catch {
      /* DI not ready yet — next toggle will retry */
    }
  }, [api, anyReactPanelOpen]);

  // Subscribe to Univer's sidebarOptions$ — close our React panels when
  // Univer opens its sidebar. The subject re-emits each time `open()` is
  // called with the latest options; the `children` field is set when a
  // panel is mounted, cleared when closed.
  const lastSeenVisibleRef = useRef(false);
  useEffect(() => {
    if (!api) return;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const injector = (api as any)._injector as
        | { get: (token: unknown) => unknown }
        | undefined;
      if (!injector) return;
      const sidebarService = injector.get('ui.sidebar.service') as
        | {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            sidebarOptions$: { subscribe: (fn: (o: any) => void) => { unsubscribe: () => void } };
            visible: boolean;
          }
        | undefined;
      if (!sidebarService) return;
      const sub = sidebarService.sidebarOptions$.subscribe((opts) => {
        const nowVisible = Boolean(opts && opts.children);
        // Only fire on rising edge — `lastSeenVisibleRef` filters out
        // the close emit our own `sidebarService.close()` triggered.
        if (nowVisible && !lastSeenVisibleRef.current) {
          ui.closeAllReactPanels();
        }
        lastSeenVisibleRef.current = nowVisible;
      });
      return () => sub.unsubscribe();
    } catch {
      /* DI not ready / service shape changed — safe to no-op */
      return;
    }
  }, [api, ui]);

  return null;
}
