import { createContext, useContext } from 'react';

/**
 * Drives the open-large-file loading overlay. ExcelJS doesn't surface
 * progress events, so we expose a *phase* enum instead of a percentage
 * — phase-aware text beats a fake progress bar that crawls to 90% and
 * sits there for half the duration.
 *
 * `null` means no overlay; an object means the overlay is open and the
 * fields describe what's happening.
 */
export type LoadingPhase = 'reading' | 'parsing' | 'mounting';

export type LoadingState = {
  fileName: string;
  /** File size in bytes, if known (xlsx open knows; co-edit seed may not). */
  sizeBytes?: number;
  phase: LoadingPhase;
  /** Unix ms when the overlay first opened — used for the elapsed timer. */
  startedAt: number;
  /** Set when the load failed. The overlay flips into an error card with
   *  this message + a Dismiss button. */
  error?: string;
  /** Optional retry handler called when the user clicks "Try again" on
   *  the error card. The overlay dismisses itself first, then invokes
   *  this — the action is then free to re-open the picker / re-fetch
   *  the seed / etc. If omitted, no Retry button is shown. */
  onRetry?: () => void;
};

export type LoadingCtxValue = {
  state: LoadingState | null;
  /** Begin / advance the overlay. Pass `null` to dismiss. Calling with a
   *  fresh state starts a new session (resets `startedAt`). Calling with
   *  only `{ phase }` shortcut updates the current session in-place. */
  set: (next: Partial<LoadingState> | null) => void;
};

export const LoadingContext = createContext<LoadingCtxValue>({
  state: null,
  set: () => undefined,
});

export function useLoading(): LoadingCtxValue {
  return useContext(LoadingContext);
}
