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
 * @casualoffice/sheets — Casual Sheets SDK
 *
 * Three surfaces:
 *   - `./signing` — anchored cell signatures (drawn / typed / uploaded).
 *   - `./embed`   — iframe postMessage protocol for host integrations.
 *   - `./sheets`  — `CasualSheets` React wrapper around Univer Sheets.
 *
 * The `./styles` side-effect entry brings in the eager plugin CSS:
 *
 *   import '@casualoffice/sheets/styles';
 */

export * from './signing';
export * from './embed';
export * from './sheets';
