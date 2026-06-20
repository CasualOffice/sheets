/**
 * Re-export shim — the replay-retry classifier moved to
 * `@casualoffice/sheets/collab` in the SDK restructure (Phase 2 step 3).
 * The app only consumes the `ReplayFailureRecord` type (driver, indicator,
 * context); the retry machinery itself now lives with the bridge in the SDK.
 * Remove once `apps/web` is a thin SDK host (Phase 3).
 */
export {
  type ReplayFailureRecord,
  type ReplayClassification,
  classifyReplayError,
} from '@casualoffice/sheets/collab';
