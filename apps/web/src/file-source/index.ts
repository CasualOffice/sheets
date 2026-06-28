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

export type {
  FileSource,
  FileSourceKind,
  OpenedWorkbook,
  RecentEntry,
  RecentId,
  SaveOptions,
  SaveResult,
} from './types';
export { FileSourceProvider, useFileSource } from './context';
export { createBrowserFileSource } from './browser-file-source';
export { createPersonalFileSource, PersonalAuthExpired } from './personal-file-source';
export { createWopiFileSource, detectWopiContext } from './wopi-file-source';
export { selectFileSource, setFileSourceKind, __resetFileSourceForTests } from './select';
export { useRecentFiles } from './useRecent';
