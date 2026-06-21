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
 * Toast notification surface for transient confirmations + dismissible
 * warnings. Designed to coexist with the persistent banners
 * (CollabDriver's OfflineBanner / ViewOnlyBanner / SelfHostBanner,
 * AutosaveRestoreBanner, PreviewBanner, LoadingOverlay) — toasts
 * are for state TRANSITIONS ("saved", "copied", "could not export"),
 * banners stay for STATE ("you are offline", "this is a preview").
 *
 * Three kinds: 'info' (neutral), 'success' (greenish), 'error' (red).
 * Each toast auto-dismisses after `duration` ms (default 3500 for
 * success/info, 6000 for error so the user has time to read).
 *
 * The hook returns shorthand methods so call sites stay short:
 *
 *   const toast = useToast();
 *   toast.success('Saved to budget.xlsx');
 *   toast.error('Could not open: file is corrupted');
 *
 * Lower-level `push` returns the toast id so callers can dismiss
 * programmatically (e.g. when a retry succeeds, dismiss the
 * "Reconnecting…" toast).
 */

export type ToastKind = 'info' | 'success' | 'error';

export interface Toast {
  id: number;
  kind: ToastKind;
  /** Body text — kept short. */
  message: string;
  /**
   * Optional action button on the toast. Use sparingly — toasts
   * are transient by design; if the user MUST take action, prefer
   * a dialog or banner.
   */
  action?: {
    label: string;
    onClick: () => void;
  };
  /** ms before auto-dismiss; <= 0 disables. */
  duration: number;
}

export type ToastInput =
  | string
  | {
      message: string;
      kind?: ToastKind;
      duration?: number;
      action?: Toast['action'];
      /**
       * Skip the ActivityContext bridge for this toast. Use when the
       * caller pushes its OWN activity entry (e.g. a retryable save
       * failure via `pushErrorWithRetry`) and a second, bare bridge
       * entry would be a duplicate. Only meaningful for `kind: 'error'`.
       */
      skipActivityLog?: boolean;
    };

export interface ToastApi {
  /** Generic push — returns the toast id. */
  push: (input: ToastInput, kind?: ToastKind) => number;
  /** `kind: 'info'` shorthand. */
  info: (message: string, opts?: { duration?: number; action?: Toast['action'] }) => number;
  /** `kind: 'success'` shorthand. */
  success: (message: string, opts?: { duration?: number; action?: Toast['action'] }) => number;
  /** `kind: 'error'` shorthand. Default duration is longer (6 s). Pass
   *  `skipActivityLog` when the caller pushes its own activity entry. */
  error: (
    message: string,
    opts?: { duration?: number; action?: Toast['action']; skipActivityLog?: boolean },
  ) => number;
  /** Dismiss by id. No-op if the toast already cleared. */
  dismiss: (id: number) => void;
}

interface ToastContextValue extends ToastApi {
  toasts: Toast[];
}

const NO_OP_API: ToastApi = {
  push: () => 0,
  info: () => 0,
  success: () => 0,
  error: () => 0,
  dismiss: () => undefined,
};

const ToastContext = createContext<ToastContextValue>({
  toasts: [],
  ...NO_OP_API,
});

const DEFAULT_DURATION: Record<ToastKind, number> = {
  info: 3500,
  success: 3500,
  error: 6000,
};

export function ToastProvider({ children }: { children: ReactNode }) {
  const [toasts, setToasts] = useState<Toast[]>([]);
  const nextIdRef = useRef(1);
  // Track auto-dismiss timers so explicit dismiss() can cancel them
  // and React strict-mode double-renders don't leak handles.
  const timersRef = useRef<Map<number, ReturnType<typeof setTimeout>>>(new Map());

  const dismiss = useCallback((id: number) => {
    const t = timersRef.current.get(id);
    if (t) {
      clearTimeout(t);
      timersRef.current.delete(id);
    }
    setToasts((prev) => prev.filter((x) => x.id !== id));
  }, []);

  const push = useCallback(
    (input: ToastInput, kindArg?: ToastKind): number => {
      const id = nextIdRef.current++;
      const normalized: Toast = (() => {
        if (typeof input === 'string') {
          const kind = kindArg ?? 'info';
          return { id, kind, message: input, duration: DEFAULT_DURATION[kind] };
        }
        const kind = input.kind ?? kindArg ?? 'info';
        return {
          id,
          kind,
          message: input.message,
          duration: input.duration ?? DEFAULT_DURATION[kind],
          action: input.action,
        };
      })();

      setToasts((prev) => [...prev, normalized]);

      // Bridge to ActivityContext for the persistent error log
      // (UX_AUDIT.md §4.1 / Phase 4 #14). Toasts vanish after 3.5–6 s;
      // the activity log keeps the failure surface around until the
      // user dismisses it. Window-event so the toast layer stays
      // unaware of the activity layer. Callers that push their own
      // activity entry (retryable failures) opt out via `skipActivityLog`.
      const skipBridge = typeof input === 'object' && input.skipActivityLog === true;
      if (normalized.kind === 'error' && !skipBridge) {
        window.dispatchEvent(
          new CustomEvent('cd:activity-error', {
            detail: { message: normalized.message },
          }),
        );
      }

      if (normalized.duration > 0) {
        const handle = setTimeout(() => dismiss(id), normalized.duration);
        timersRef.current.set(id, handle);
      }
      return id;
    },
    [dismiss],
  );

  const info = useCallback(
    (message: string, opts?: { duration?: number; action?: Toast['action'] }) =>
      push({ message, kind: 'info', ...opts }),
    [push],
  );
  const success = useCallback(
    (message: string, opts?: { duration?: number; action?: Toast['action'] }) =>
      push({ message, kind: 'success', ...opts }),
    [push],
  );
  const error = useCallback(
    (
      message: string,
      opts?: { duration?: number; action?: Toast['action']; skipActivityLog?: boolean },
    ) => push({ message, kind: 'error', ...opts }),
    [push],
  );

  // Drain timers on unmount so HMR / tab-close doesn't leak handles.
  useEffect(() => {
    const timers = timersRef.current;
    return () => {
      for (const t of timers.values()) clearTimeout(t);
      timers.clear();
    };
  }, []);

  const value = useMemo<ToastContextValue>(
    () => ({ toasts, push, info, success, error, dismiss }),
    [toasts, push, info, success, error, dismiss],
  );

  return <ToastContext.Provider value={value}>{children}</ToastContext.Provider>;
}

/**
 * Read the toast API. Safe to call outside a provider — returns
 * no-op methods so call sites in tests / unmounted islands don't
 * have to defensively guard.
 */
export function useToast(): ToastApi {
  const ctx = useContext(ToastContext);
  return ctx;
}

/** Read the current toast list. Used by `<ToastContainer />`. */
export function useToastList(): Toast[] {
  return useContext(ToastContext).toasts;
}
