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
 * mountCasualSheets — the framework-agnostic imperative mount (doc 38 §1).
 *
 * React hosts render `<CasualSheets {...config} />`; vanilla / non-React hosts
 * call `mountCasualSheets(container, config)` and get back the same
 * `CasualSheetsAPI` the component hands to `onReady`, plus a `destroy()` that
 * unmounts the React root. This is the sheets peer of the docs SDK's
 * `renderAsync(input, container, options)` — one programmatic entry point that
 * returns the full imperative handle, so a host never has to stand up its own
 * React root to drive the editor.
 *
 * (`mountEmbedded` in `../embed-runtime` is a DIFFERENT, iframe-only entry: it
 * boots the postMessage embed runtime and returns `void`. This one mounts the
 * live editor into a DOM node the host already controls and returns the API.)
 *
 * The config is exactly `CasualSheetsProps` — the same declarative object the
 * component takes — so the two mount styles stay key-for-key identical (doc 38
 * §1). `react` / `react-dom` are peer deps the host already provides.
 */

import { createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { CasualSheets, type CasualSheetsProps } from './CasualSheets';
import type { CasualSheetsAPI } from './api';

/** The handle returned by {@link mountCasualSheets}. */
export interface MountedCasualSheets {
  /** The imperative editor API — identical to the ref handed to `onReady`
   *  (snapshot I/O, xlsx import/export, selection, command dispatch, undo/redo,
   *  the `.on()/.off()` emitter, document mode, `api.univer` escape hatch). */
  readonly api: CasualSheetsAPI;
  /** Unmount the React root and dispose the editor. The imperative-mount peer
   *  of a React unmount (doc 38 §4 `destroy`). Idempotent. */
  destroy(): void;
}

/**
 * Mount `<CasualSheets>` into a DOM node without a host React tree, and resolve
 * once the editor is ready with its {@link CasualSheetsAPI} + a `destroy()`.
 *
 * Any `onReady` / `onError` you pass in `options` still fires — this composes
 * with them rather than replacing them. The returned promise resolves on the
 * first `ready`; if boot fails before ready, it rejects with the boot error
 * (so callers awaiting the mount never hang). Errors after ready reach your
 * `onError` (prop or `api.on('error', …)`), not the promise.
 */
export function mountCasualSheets(
  container: HTMLElement,
  options: CasualSheetsProps,
): Promise<MountedCasualSheets> {
  const root: Root = createRoot(container);
  let destroyed = false;
  const destroy = () => {
    if (destroyed) return;
    destroyed = true;
    // Defer off any in-flight render commit (Univer owns its own React root and
    // warns on a synchronous unmount mid-render), mirroring the component's
    // microtask-deferred disposal.
    queueMicrotask(() => root.unmount());
  };

  return new Promise<MountedCasualSheets>((resolve, reject) => {
    let settled = false;
    const onReady = (api: CasualSheetsAPI) => {
      // Preserve the host's own onReady first, then settle the mount promise.
      options.onReady?.(api);
      if (settled) return;
      settled = true;
      resolve({ api, destroy });
    };
    const onError = (err: Error) => {
      options.onError?.(err);
      // A boot error before `ready` would otherwise leave the promise pending.
      if (settled) return;
      settled = true;
      reject(err);
    };
    try {
      root.render(createElement(CasualSheets, { ...options, onReady, onError }));
    } catch (err) {
      destroy();
      reject(err instanceof Error ? err : new Error(String(err)));
    }
  });
}
