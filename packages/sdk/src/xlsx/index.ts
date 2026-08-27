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
 * xlsx — `.xlsx` ↔ Univer IWorkbookData converters.
 *
 * Both directions:
 *
 *   import { xlsxToWorkbookData, workbookDataToXlsx } from '@casualoffice/sheets/xlsx';
 *
 * Browser main threads run the converter in a Web Worker so multi-MB
 * workbooks don't block the UI. Node and existing Worker contexts reuse the
 * same converter implementation directly and require no DOM or Worker global.
 *
 * Fidelity scope:
 *   - Values + formulas
 *   - Font (family, size, bold, italic, underline, color)
 *   - Fill (solid background)
 *   - Alignment (horizontal, vertical, wrap)
 *   - Number format
 *   - Borders (thin, per side, color preserved)
 *   - Merges
 *   - Sheet order + names
 *   - Tables, comments, data validation, page setup, named ranges (resources)
 *
 * Opaque resources are carried through the hidden resources sidecar; raw
 * drawings, pivots, macros, external links, and threaded comments additionally
 * round-trip through OOXML passthrough payloads.
 *
 * The shared utilities below (style mappers + resource readers) are
 * exposed for hosts that ship their own xlsx export path and want to
 * stay in lockstep with this importer's shape. Casual Sheets' own
 * apps/web uses them for that reason. Other consumers can ignore.
 */

export { xlsxToWorkbookData, type ImportedWorkbook } from './import';
export { workbookDataToXlsx, type ExportExtras } from './export';
export * from './style-mapping';
export * from './constants';
export * from './comments-resource';
export * from './page-setup-resource';
export * from './data-validation-resource';
export * from './tables-resource';
export * from './passthrough-resource';
// Native pivot generation (in-app pivots → real xl/pivotTables) + the
// injection helper, so app hosts can compose native pivots into an export.
export { generateNativePivot, type NativePivotInput, type PivotAgg } from './pivot-export';
export { applyPivotsToZip, type PivotPassthroughPayload } from './pivot-passthrough';
