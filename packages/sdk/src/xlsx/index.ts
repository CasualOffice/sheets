/**
 * xlsx — `.xlsx` ↔ Univer IWorkbookData converters.
 *
 * Phase-A surface (this release): IMPORT only.
 *
 *   import { xlsxToWorkbookData } from '@schnsrw/casual-sheets/xlsx';
 *
 * The parser runs in a Web Worker so multi-MB workbooks don't block
 * the main thread; bundler must support the `new Worker(new URL(...),
 * import.meta.url)` pattern (Vite, esbuild's bundler, modern webpack
 * with worker-plugin).
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
 * Accepts loss: charts, drawings, pivots, sparklines, advanced borders
 * (dashed/double), themes.
 *
 * The shared utilities below (style mappers + resource readers) are
 * exposed for hosts that ship their own xlsx export path and want to
 * stay in lockstep with this importer's shape. Casual Sheets' own
 * apps/web uses them for that reason. Other consumers can ignore.
 */

export { xlsxToWorkbookData, type ImportedWorkbook } from './import';
export * from './style-mapping';
export * from './constants';
export * from './comments-resource';
export * from './page-setup-resource';
export * from './data-validation-resource';
export * from './tables-resource';
export * from './passthrough-resource';
