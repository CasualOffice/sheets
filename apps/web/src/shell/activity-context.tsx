import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useRef,
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
 * v1 shipped the log itself (icon + badge + popover listing the last
 * `MAX_ENTRIES`). v2 wires per-entry retry handlers: call sites that
 * own an idempotent recovery action (save, export, restore) push via
 * `pushErrorWithRetry`, and the popover renders a Retry button that
 * re-runs the stored closure. Retry callbacks live in a ref-held Map
 * keyed by entry id — NOT on the entry object, which stays
 * JSON-serializable (it travels the window-event bridge below).
 *
 * The toast → activity bridge is a window event (`cd:activity-error`)
 * so ToastContext stays unaware of ActivityContext. Any other system
 * (background autosave, future server-push) can fire the same event
 * to enter the log.
 */
export const ACTIVITY_EVENT = 'cd:activity-error';

/** Coarse action class for an entry — drives copy / grouping only.
 *  The actual recovery closure lives in the provider's ref map. */
export type ActivityActionKind = 'save' | 'export' | 'restore' | 'other';

export interface ActivityEntry {
  id: number;
  kind: 'error';
  message: string;
  timestamp: number;
  /** What kind of action failed (when known). Serializable — the
   *  retry closure itself is held separately so the entry can travel
   *  the window-event bridge / be logged. */
  actionKind?: ActivityActionKind;
}

/** A retry handler. Must re-capture fresh state on each call (e.g.
 *  re-snapshot the workbook) — never close over stale data. Resolves
 *  on success; rejects to keep the entry and surface the new error. */
export type RetryFn = () => Promise<void>;

export interface ActivityCtx {
  entries: ActivityEntry[];
  /** Unread count — drives the badge. */
  unread: number;
  /** Programmatic push (used by tests; production callers use the event). */
  pushError: (message: string) => void;
  /** Push an error that carries a retry handler. Returns the entry id. */
  pushErrorWithRetry: (message: string, retry: RetryFn, actionKind?: ActivityActionKind) => number;
  /** Whether an entry has a stored retry handler (drives the button). */
  hasRetry: (id: number) => boolean;
  /** Whether a retry is currently in flight for an entry. */
  isRetrying: (id: number) => boolean;
  /** Run the stored retry for an entry. On success the entry is
   *  dismissed; on failure it's kept and the rejection bubbles to a
   *  fresh error entry. No-op if the entry has no handler. */
  retryEntry: (id: number) => Promise<void>;
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
  pushErrorWithRetry: () => 0,
  hasRetry: () => false,
  isRetrying: () => false,
  retryEntry: async () => undefined,
  markAllRead: () => undefined,
  dismiss: () => undefined,
  clearAll: () => undefined,
};

const ActivityContext = createContext<ActivityCtx | null>(null);

const MAX_ENTRIES = 25;

export function ActivityProvider({ children }: { children: ReactNode }) {
  const [entries, setEntries] = useState<ActivityEntry[]>([]);
  const [unread, setUnread] = useState(0);
  // In-flight retry ids — drives the per-row spinner / disabled state.
  const [retrying, setRetrying] = useState<ReadonlySet<number>>(() => new Set());

  // Monotonic id source. Must never reuse a value, or a stale entry in
  // the retry map could be invoked for a new entry that happens to land
  // on the same id (the old `prev[0].id + 1` scheme collided after a
  // dismiss-then-push).
  const nextIdRef = useRef(1);
  // Retry closures, keyed by entry id. Held in a ref (not on the entry)
  // so `ActivityEntry` stays JSON-serializable. Pruned on dismiss /
  // clearAll / successful retry so it can't leak callbacks past the cap.
  const retryMapRef = useRef<Map<number, RetryFn>>(new Map());

  const insertEntry = useCallback((message: string, actionKind?: ActivityActionKind): number => {
    const id = nextIdRef.current++;
    const next: ActivityEntry = {
      id,
      kind: 'error',
      message,
      timestamp: Date.now(),
      ...(actionKind ? { actionKind } : {}),
    };
    setEntries((prev) => {
      // Newest-first, capped at MAX_ENTRIES so a runaway error loop
      // can't OOM the browser. Older entries fall off the tail — drop
      // their retry closures too.
      const merged = [next, ...prev];
      if (merged.length > MAX_ENTRIES) {
        for (const dropped of merged.slice(MAX_ENTRIES)) {
          retryMapRef.current.delete(dropped.id);
        }
      }
      return merged.slice(0, MAX_ENTRIES);
    });
    setUnread((n) => n + 1);
    return id;
  }, []);

  const pushError = useCallback(
    (message: string) => {
      insertEntry(message);
    },
    [insertEntry],
  );

  const pushErrorWithRetry = useCallback(
    (message: string, retry: RetryFn, actionKind: ActivityActionKind = 'other'): number => {
      const id = insertEntry(message, actionKind);
      retryMapRef.current.set(id, retry);
      return id;
    },
    [insertEntry],
  );

  const hasRetry = useCallback((id: number) => retryMapRef.current.has(id), []);
  const isRetrying = useCallback((id: number) => retrying.has(id), [retrying]);

  const removeEntry = useCallback((id: number) => {
    retryMapRef.current.delete(id);
    setEntries((prev) => prev.filter((e) => e.id !== id));
    setRetrying((prev) => {
      if (!prev.has(id)) return prev;
      const next = new Set(prev);
      next.delete(id);
      return next;
    });
  }, []);

  const retryEntry = useCallback(
    async (id: number) => {
      const fn = retryMapRef.current.get(id);
      if (!fn) return;
      // Guard against double-fire (rapid clicks / Enter-repeat).
      let already = false;
      setRetrying((prev) => {
        if (prev.has(id)) {
          already = true;
          return prev;
        }
        const next = new Set(prev);
        next.add(id);
        return next;
      });
      if (already) return;
      try {
        await fn();
        // Success — the failure is resolved, so the entry goes away.
        removeEntry(id);
      } catch (err) {
        // Keep the entry (still actionable) and surface the new error.
        // The failed entry stays retryable so the user can try again.
        setRetrying((prev) => {
          const next = new Set(prev);
          next.delete(id);
          return next;
        });
        const msg = err instanceof Error ? err.message : String(err);
        insertEntry(`Retry failed: ${msg}`);
      }
    },
    [insertEntry, removeEntry],
  );

  // Bridge from ToastContext (window event). Decouples the two
  // contexts so the toast layer can stay generic. Entries from the
  // bridge carry no retry handler — only call sites that own a real
  // recovery action use `pushErrorWithRetry` directly.
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
  const dismiss = useCallback((id: number) => removeEntry(id), [removeEntry]);
  const clearAll = useCallback(() => {
    retryMapRef.current.clear();
    setEntries([]);
    setUnread(0);
    setRetrying(new Set());
  }, []);

  const value = useMemo<ActivityCtx>(
    () => ({
      entries,
      unread,
      pushError,
      pushErrorWithRetry,
      hasRetry,
      isRetrying,
      retryEntry,
      markAllRead,
      dismiss,
      clearAll,
    }),
    [
      entries,
      unread,
      pushError,
      pushErrorWithRetry,
      hasRetry,
      isRetrying,
      retryEntry,
      markAllRead,
      dismiss,
      clearAll,
    ],
  );

  // DEV-only test seam: lets e2e specs push an entry WITH a retry
  // handler (the window-event bridge can only carry serializable data,
  // so it can't deliver a closure). Mirrors the `__toastLog__` seam in
  // file-actions. Tree-shaken from production builds.
  useEffect(() => {
    if (!import.meta.env.DEV) return;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__activityRetry__ = {
      push: (message: string, retry: RetryFn, actionKind?: ActivityActionKind) =>
        pushErrorWithRetry(message, retry, actionKind),
    };
    return () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      delete (window as any).__activityRetry__;
    };
  }, [pushErrorWithRetry]);

  return <ActivityContext.Provider value={value}>{children}</ActivityContext.Provider>;
}

export function useActivity(): ActivityCtx {
  return useContext(ActivityContext) ?? FALLBACK;
}
