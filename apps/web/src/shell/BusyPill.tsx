import { useBusy } from '../busy-context';

/**
 * Title-bar pill that surfaces the BusyContext state. Tiny by design:
 * a spinner + the current label, no dialog, no backdrop. Click-through
 * (pointer-events: none) so the user can still hit other UI; the slow
 * action will resolve on its own when the main thread frees up.
 */
export function BusyPill() {
  const { state } = useBusy();
  if (!state) return null;
  return (
    <span
      className="busy-pill"
      data-testid="busy-pill"
      role="status"
      aria-live="polite"
    >
      <span className="busy-pill__spinner" aria-hidden="true" />
      <span className="busy-pill__label">{state.label}</span>
    </span>
  );
}
