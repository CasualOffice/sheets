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
 * Name of a hidden worksheet we use to stash JSON we can't represent
 * natively in xlsx (Univer plugin state — e.g. table definitions, outline
 * groups). On open we recognize and consume it, never showing it to the
 * user. Defined in its own module so both the parser worker and the
 * exporter worker can import it without dragging the other one's
 * ExcelJS code into their bundle.
 */
export const RESOURCES_SHEET = '__casual_sheets_resources__';
