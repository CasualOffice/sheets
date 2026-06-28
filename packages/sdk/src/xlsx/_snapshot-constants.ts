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
 * Snapshot defaults pulled out of apps/web/src/snapshot.ts so the
 * SDK's xlsx parser doesn't have to import that whole module. Keep
 * these in sync with the host snapshot module if the defaults drift.
 *
 * UNIVER_VERSION must match the runtime Univer the host boots — the
 * appVersion field on the IWorkbookData snapshot is checked at unit
 * mount and a mismatch warns in dev. Sheet apps' `../snapshot` reads
 * the version from the workspace's @univerjs/core dep; we hardcode
 * the same minor here because the SDK declares @univerjs/* as
 * `^0.24.0` peer.
 */
export const INITIAL_ROWS = 1024;
export const INITIAL_COLUMNS = 26;
export const UNIVER_VERSION = '0.24.0';
