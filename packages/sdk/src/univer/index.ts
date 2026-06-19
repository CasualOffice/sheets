/**
 * Internal Univer wiring shared by the SDK editor and the host app.
 *
 * Phase 1 (Batch 1) of the SDK migration lifts the lazy plugin loader out of
 * `apps/web` so the editor can eventually run entirely from `@casualoffice/sheets`.
 * It's pure Univer/DI code (no React context, FileSource, collab, or routing) with
 * module-level singleton state, so the host and the SDK must resolve this one
 * module instance. Later batches add the facade-coupled helpers (dev-helpers,
 * paste-merge-hook, zoom override) + the editor core. See
 * `docs/SDK_MIGRATION_PIPELINE.md` Phase 1.
 *
 * @internal — not part of the SDK's semver surface; consumers use `<CasualSheets>`.
 */

// Lazy plugin loader (CF, DV, hyperlink, note, table, thread-comment, drawing,
// sort, filter, find-replace) + the module-level Univer holder.
export * from './lazy-plugins';
