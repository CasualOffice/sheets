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

import { createBrowserFileSource } from './browser-file-source';
import { createPersonalFileSource } from './personal-file-source';
import { createWopiFileSource, detectWopiContext } from './wopi-file-source';
import type { FileSource, FileSourceKind } from './types';

/**
 * Per-deploy `FileSource` picker.
 *
 * Selection is driven by:
 *   - URL — when the page is loaded with an access_token query, the
 *     visitor is in an embedded-host context (WOPI / Mode 2). The
 *     source is bound to that single file id, and the personal auth
 *     gate is skipped entirely.
 *   - React `AuthProvider` — when the personal-mode auth probe lands
 *     `authenticated`, the provider calls
 *     `setFileSourceKind('personal')`; when it lands `disabled` /
 *     `unauthenticated` / `loading`, it stays on the browser source.
 *
 * Pre-React call sites in `file-actions` consume the same cached
 * instance via `selectFileSource()`. One cached instance per
 * `kind` keeps the React provider, the non-React file-actions
 * saves, and the live-feed subscribers pointed at a single object.
 *
 * Phase B: always browser.
 * Phase C: browser by default; provider swaps to personal once the
 *          auth gate reports authenticated.
 * Phase D: wopi when the URL carries a token — wins over personal
 *          so an embedded host doesn't have to know about the
 *          self-host auth surface.
 */

type Kind = FileSourceKind;

const cache = new Map<Kind, FileSource>();
let activeKind: Kind = 'browser';

function makeFor(kind: Kind): FileSource {
  if (kind === 'wopi') {
    const ctx = detectWopiContext();
    if (ctx) return createWopiFileSource(ctx);
    // Boot probe didn't find a token after all — degrade gracefully.
    return createBrowserFileSource();
  }
  if (kind === 'personal') return createPersonalFileSource();
  return createBrowserFileSource();
}

/** One-shot boot probe — runs once at module load and returns the
 *  initial active kind so the React provider doesn't have to
 *  duplicate the URL check. */
function initialKind(): Kind {
  if (detectWopiContext()) return 'wopi';
  return 'browser';
}

// Initialise active kind eagerly so `selectFileSource()` (called
// from non-React file-actions before the provider mounts) sees the
// right value on the first invocation.
activeKind = initialKind();

export function selectFileSource(): FileSource {
  let src = cache.get(activeKind);
  if (!src) {
    src = makeFor(activeKind);
    cache.set(activeKind, src);
  }
  return src;
}

/** Swap the active source. Called from `AuthProvider` when the
 *  user's auth state crosses an authenticated boundary. */
export function setFileSourceKind(kind: Kind): FileSource {
  activeKind = kind;
  return selectFileSource();
}

/** Test hook — drop every cached instance so a fresh probe runs
 *  next time. Production code should not call this. */
export function __resetFileSourceForTests(): void {
  cache.clear();
  activeKind = 'browser';
}
