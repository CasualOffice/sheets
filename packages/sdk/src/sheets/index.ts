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
 * Sheets surface — React wrappers around Univer Sheets.
 */
export { CasualSheets, type CasualSheetsProps } from './CasualSheets';
export {
  createCasualSheetsAPI,
  type CasualSheetsAPI,
  type DocumentMode,
  type RangeRef,
  type CommandRecord,
} from './api';
export { applyReadOnly, applyCommentOnly, getEditable } from './read-only';
export {
  setMentionProvider,
  getMentionProvider,
  filterMentionCandidates,
  type MentionCandidate,
  type MentionProvider,
} from './mention-source';
export { CasualMentionIOService } from './mention-io';

// AI surface — the `ai` prop contract on `CasualSheets`, its transport type,
// and the shipped transport implementations (direct / collab single-round /
// desktop-native + the `createSheetsAiTransport` factory).
export {
  AiPanelSurface,
  DirectAiTransport,
  CollabAiTransport,
  DesktopAiTransport,
  createSheetsAiTransport,
  type SheetsAiConfig,
  type SheetsAiAction,
  type SheetsAiRenderContext,
  type SheetsAiTransport,
  type SheetsAiLlmPayload,
  type SheetsAiLlmResult,
  type SheetsAiToolExecutor,
  type CreateSheetsAiTransportOptions,
} from '../ai';
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
