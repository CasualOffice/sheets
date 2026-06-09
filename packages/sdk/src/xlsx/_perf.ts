/**
 * Tiny perf shim — drop-in replacement for apps/web's `../perf` so the
 * SDK doesn't have to pull in a User Timing dependency. The host can
 * wrap individual SDK calls with its own timing if it wants spans;
 * inside the SDK they're cheap no-op forwards.
 */

export function timeIt<T>(_label: string, fn: () => T): T {
  return fn();
}

export async function timeItAsync<T>(_label: string, fn: () => Promise<T>): Promise<T> {
  return fn();
}
