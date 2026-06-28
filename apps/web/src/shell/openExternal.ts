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
 * Open an external link in the user's default browser.
 *
 * Inside the Casual Office desktop shell, `window.open()` silently no-ops in the
 * Tauri webview (so Help / "View on GitHub" / "Report a bug" did nothing). Route
 * through the shell's scheme-guarded `open_external` command when the Tauri
 * bridge is present; fall back to `window.open` on the web.
 */
export function openExternal(url: string): void {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const invoke = (window as any)?.__TAURI__?.core?.invoke as
    | ((cmd: string, args?: unknown) => Promise<unknown>)
    | undefined;
  if (invoke) {
    void invoke('open_external', { url }).catch(() => {
      window.open(url, '_blank', 'noopener,noreferrer');
    });
    return;
  }
  window.open(url, '_blank', 'noopener,noreferrer');
}
