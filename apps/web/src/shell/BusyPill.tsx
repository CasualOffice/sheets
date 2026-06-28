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
