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
