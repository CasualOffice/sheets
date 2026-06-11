import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
  type ReactNode,
} from 'react';

/**
 * Activity log — UX_AUDIT.md §4.1 / Phase 4 #14. Toasts are
 * transient (3.5–6 s) and gone; a user who looks away during a
 * failed save has no way to know it happened. The activity log is
 * the persistent surface that keeps recent errors around until the
 * user dismisses them.
 *
 * v1 ships the log itself (icon + badge + popover listing the last
 * `MAX_ENTRIES`). v2 wires per-entry retry handlers — that needs
 * per-call-site action context that doesn't exist at the toast
 * layer today, so it's deferred to a follow-up.
 *
 * The toast → activity bridge is a window event (`cd:activity-error`)
 * so ToastContext stays unaware of ActivityContext. Any other system
 * (background autosave, future server-push) can fire the same event
 * to enter the log.
 */
export const ACTIVITY_EVENT = 'cd:activity-error';

export interface ActivityEntry {
  id: number;
  kind: 'error';
  message: string;
  timestamp: number;
}

export interface ActivityCtx {
  entries: ActivityEntry[];
  /** Unread count — drives the badge. */
  unread: number;
  /** Programmatic push (used by tests; production callers use the event). */
  pushError: (message: string) => void;
  /** Clear the unread badge — called when the popover opens. */
  markAllRead: () => void;
  /** Remove a single entry. */
  dismiss: (id: number) => void;
  /** Wipe the log. */
  clearAll: () => void;
}

const FALLBACK: ActivityCtx = {
  entries: [],
  unread: 0,
  pushError: () => undefined,
  markAllRead: () => undefined,
  dismiss: () => undefined,
  clearAll: () => undefined,
};

const ActivityContext = createContext<ActivityCtx | null>(null);

const MAX_ENTRIES = 25;

export function ActivityProvider({ children }: { children: ReactNode }) {
  const [entries, setEntries] = useState<ActivityEntry[]>([]);
  const [unread, setUnread] = useState(0);

  const pushError = useCallback((message: string) => {
    setEntries((prev) => {
      const next: ActivityEntry = {
        id: prev.length ? prev[0].id + 1 : 1,
        kind: 'error',
        message,
        timestamp: Date.now(),
      };
      // Newest-first, capped at MAX_ENTRIES so a runaway error loop
      // can't OOM the browser. Older entries fall off the tail.
      return [next, ...prev].slice(0, MAX_ENTRIES);
    });
    setUnread((n) => n + 1);
  }, []);

  // Bridge from ToastContext (window event). Decouples the two
  // contexts so the toast layer can stay generic.
  useEffect(() => {
    const onErr = (e: Event) => {
      const detail = (e as CustomEvent<{ message?: string }>).detail;
      if (!detail || typeof detail.message !== 'string') return;
      pushError(detail.message);
    };
    window.addEventListener(ACTIVITY_EVENT, onErr as EventListener);
    return () => window.removeEventListener(ACTIVITY_EVENT, onErr as EventListener);
  }, [pushError]);

  const markAllRead = useCallback(() => setUnread(0), []);
  const dismiss = useCallback((id: number) => {
    setEntries((prev) => prev.filter((e) => e.id !== id));
  }, []);
  const clearAll = useCallback(() => {
    setEntries([]);
    setUnread(0);
  }, []);

  const value = useMemo<ActivityCtx>(
    () => ({ entries, unread, pushError, markAllRead, dismiss, clearAll }),
    [entries, unread, pushError, markAllRead, dismiss, clearAll],
  );

  return <ActivityContext.Provider value={value}>{children}</ActivityContext.Provider>;
}

export function useActivity(): ActivityCtx {
  return useContext(ActivityContext) ?? FALLBACK;
}
