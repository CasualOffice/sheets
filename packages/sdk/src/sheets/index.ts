/**
 * Sheets surface — React wrappers around Univer Sheets.
 */
export { CasualSheets, type CasualSheetsProps } from './CasualSheets';
export { createCasualSheetsAPI, type CasualSheetsAPI, type RangeRef } from './api';
export { applyReadOnly, getEditable } from './read-only';
export {
  CasualSheetsIframe,
  type CasualSheetsIframeProps,
  type CasualSheetsIframeRef,
  type HostFileBridge,
} from './CasualSheetsIframe';

// Chrome extension API types — type-only re-export so hosts can type their
// `extensions` prop straight off `@casualoffice/sheets` (the values/components
// live in the `@casualoffice/sheets/chrome` subpath, code-split there).
export type {
  ChromeExtensions,
  ToolbarExtension,
  MenuExtension,
  PanelExtension,
  DialogExtension,
  DialogComponentProps,
  PanelComponentProps,
  MenuTarget,
  DialogKind,
} from '../chrome';
