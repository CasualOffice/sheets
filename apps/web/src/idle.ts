/**
 * Run heavy, non-urgent work at the next idle moment so it never hitches the
 * UI mid-interaction. Used by the background snapshot loops (autosave,
 * version-history) whose `wb.save()` deep clone can take hundreds of ms on a
 * large workbook — running it on a timer would freeze the grid mid-keystroke.
 *
 * `timeout` guarantees the callback still runs promptly even if the page never
 * goes idle (continuous typing). Falls back to a near-immediate `setTimeout`
 * where `requestIdleCallback` is unavailable (older Safari).
 */

export interface IdleHandle {
  cancel: () => void;
}

export function runWhenIdle(cb: () => void, timeout = 2_000): IdleHandle {
  const w = window as typeof window & {
    requestIdleCallback?: (cb: () => void, opts?: { timeout: number }) => number;
    cancelIdleCallback?: (id: number) => void;
  };
  if (typeof w.requestIdleCallback === 'function') {
    const id = w.requestIdleCallback(cb, { timeout });
    return { cancel: () => w.cancelIdleCallback?.(id) };
  }
  const id = window.setTimeout(cb, 1);
  return { cancel: () => window.clearTimeout(id) };
}
