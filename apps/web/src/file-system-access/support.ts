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
 * Feature detection for the File System Access API. Chromium-only at
 * the time of writing (Chrome / Edge / Brave / Opera on desktop). All
 * call sites must guard with `isFsaSupported()` and fall back to the
 * download-blob path when false — Firefox / Safari users still get
 * Save, they just get the existing browser download UX.
 */

export function isFsaSupported(): boolean {
  return typeof window !== 'undefined' && 'showDirectoryPicker' in window;
}
