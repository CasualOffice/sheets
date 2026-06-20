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
  type CollabHandle,
  type CollabRole,
  type CollabConnectionStatus,
} from './attachCollab';

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
