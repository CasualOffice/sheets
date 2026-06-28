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

import { createContext, useContext } from 'react';
import type * as Y from 'yjs';
import type { ReplayFailureRecord } from '@casualoffice/sheets/collab';

export type CollabStatus = 'off' | 'connecting' | 'live' | 'offline' | 'denied';
// Share link-roles (sharing-model §6). `comment` is the middle tier: cells are
// read-only but threaded comments stay usable (enforced via applyCommentOnly).
export type CollabRole = 'view' | 'comment' | 'write';
/**
 * `in-sync`   — local Y.Doc state vector matches every visible peer.
 *               Common steady state.
 * `syncing`   — one or more peers disagree, but the disagreement is
 *               fresh (< 15 s). Normal during active editing.
 * `diverged`  — disagreement has persisted past the grace window.
 *               Surfaces as a warning pill so the user knows their
 *               edits and their peers' edits may not match. Refresh
 *               usually recovers.
 */
export type SyncHealth = 'in-sync' | 'syncing' | 'diverged';

export type CollabCtxValue = {
  /** True when the active build was made with VITE_COLLAB_ENABLED. */
  enabled: boolean;
  /** Current room id, or null when not in a room. */
  roomId: string | null;
  /** WebSocket transport status. `off` when no room is joined; `denied`
   *  when the server rejected the password / auth. */
  status: CollabStatus;
  /** Effective role for this client in the active room. Defaults to
   *  `write` outside a room (the local single-user editor). */
  role: CollabRole;
  /** Aggregate Yjs sync health across visible peers. Only meaningful
   *  when `status === 'live'`; falls back to `in-sync` otherwise. */
  syncHealth: SyncHealth;
  /** Number of OTHER peers currently visible via awareness (i.e. the
   *  presence count NOT including the local user). Lets the indicator
   *  show "Live · 2 peers" instead of just "Live". Zero outside a
   *  room and during connect. */
  peerCount: number;
  /** Number of locally-emitted mutation records added to the op log
   *  since the WebSocket dropped to `offline`. Resets to 0 the moment
   *  the connection returns to `live`. Lets the offline banner /
   *  indicator show "3 changes queued" so users can see their edits
   *  aren't being silently lost. */
  queuedLocal: number;
  /** Number of REMOTE mutations that have thrown during local replay
   *  since the bridge started. Each one is a candidate divergence
   *  (peer's edit didn't apply here → state vectors will disagree).
   *  Surfaced as a warning pill so the user sees the failure before
   *  acting on a stale view. Refresh re-fetches the full snapshot
   *  and typically recovers. Sticky for the session (doesn't reset). */
  replayFailures: number;
  /** Per-mutation dead-letter records for the latest failures (cap
   *  20). Surfaces in the CollabIndicator's expand-on-click panel so
   *  the user can see WHICH mutations failed, not just a count.
   *  Empty array outside a room and on a healthy session. */
  replayDeadLetter: readonly ReplayFailureRecord[];
  /** Active room's Yjs document — null when not in a room. Exposed
   *  via context so the HistoryPanel can subscribe to the op-log
   *  array without piping it through three layers of props. */
  doc: Y.Doc | null;
};

export const CollabContext = createContext<CollabCtxValue>({
  enabled: false,
  roomId: null,
  status: 'off',
  role: 'write',
  syncHealth: 'in-sync',
  peerCount: 0,
  queuedLocal: 0,
  replayFailures: 0,
  replayDeadLetter: [],
  doc: null,
});

export function useCollab(): CollabCtxValue {
  return useContext(CollabContext);
}
