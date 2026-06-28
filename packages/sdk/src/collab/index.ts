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
 * `@casualoffice/sheets/collab` — opt-in real-time co-editing.
 *
 * The editor is collab-unaware until a host calls `attachCollab(api, opts)`.
 * Yjs + Hocuspocus is the realtime transport; authoritative persistence stays
 * host-side (WOPI / backend) via the save/exit event contract.
 *
 * Requires the host to provide `yjs` and `@hocuspocus/provider` (peer deps) so
 * there's a single Yjs instance in the graph — two copies break `Y.Doc`
 * identity and awareness.
 */

export {
  attachCollab,
  type AttachCollabOptions,
  type CollabAttachable,
  type CollabHandle,
  type CollabRole,
  type CollabConnectionStatus,
} from './attachCollab';
// Secure share-link capability shape (sharing-model §6.1) — referenced by
// AttachCollabOptions.share.
export { type WsUrlShare } from './ws-url';

// The mutation bridge — the framework-agnostic core (subscribes to
// onMutationExecutedForCollab, replays with fromCollab, guards __splitChunk__).
// Exposed for hosts that drive their own provider/doc lifecycle instead of
// using attachCollab's batteries-included transport.
export {
  startBridge,
  type BridgeHandle,
  type BridgeOptions,
  SYNCED_MUTATIONS,
  REVERTABLE_MUTATIONS,
} from './bridge';
export {
  type ReplayFailureRecord,
  type ReplayClassification,
  classifyReplayError,
} from './replay-retry';
export { deepRewriteUnitId, rewriteJson1OpPathUnitId } from './bridge-helpers';
