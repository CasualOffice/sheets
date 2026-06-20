/**
 * Re-export shim — the mutation bridge moved to
 * `@casualoffice/sheets/collab` in the SDK restructure (Phase 2 step 3).
 * Kept so the app's `CollabDriver` import stays stable; remove once
 * `apps/web` becomes a thin SDK host that calls `attachCollab` directly
 * (Phase 3).
 */
export {
  startBridge,
  type BridgeHandle,
  type BridgeOptions,
  SYNCED_MUTATIONS,
  REVERTABLE_MUTATIONS,
} from '@casualoffice/sheets/collab';
