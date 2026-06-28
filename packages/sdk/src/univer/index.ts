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
