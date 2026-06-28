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

import { createContext, useContext, useEffect, useState, type ReactNode } from 'react';
import { useAuth } from '../auth';
import { createBrowserFileSource } from './browser-file-source';
import { selectFileSource, setFileSourceKind } from './select';
import { detectWopiContext } from './wopi-file-source';
import type { FileSource } from './types';

/**
 * React surface for the active `FileSource`. The provider listens
 * to the auth state and swaps the source between the cached
 * browser/personal instances as the user signs in / out.
 *
 * Phase B: the provider always rendered the browser source.
 * Phase C: when AuthState.kind === 'authenticated', the provider
 *          flips to the personal source so /files/* HTTP calls
 *          replace IDB reads. Sign-out flips back to browser
 *          (recent files in IDB are preserved across modes since
 *          the cache holds one instance per kind).
 *
 * Provider is mounted INSIDE `AuthProvider` and OUTSIDE every UI
 * consumer so the File menu, the home screen, and `file-actions`
 * share the same source instance.
 */

const FileSourceContext = createContext<FileSource | null>(null);

export function FileSourceProvider({ children }: { children: ReactNode }) {
  const { state } = useAuth();
  // `state.kind` drives which cached source the consumers see.
  // Keep `source` in React state so a swap triggers a re-render of
  // every `useFileSource()` consumer.
  const [source, setSource] = useState<FileSource>(() => selectFileSource());

  useEffect(() => {
    // URL-token boot path wins over the AuthState. An embedded host
    // is authenticated against its own identity (the JWT), not the
    // personal-mode users table, so the AuthProvider's 'disabled' /
    // 'unauthenticated' / 'authenticated' verdicts don't apply here.
    if (detectWopiContext()) {
      setSource(setFileSourceKind('wopi'));
      return;
    }
    if (state.kind === 'authenticated') {
      setSource(setFileSourceKind('personal'));
    } else {
      // 'disabled' (mode=none / Pages), 'unauthenticated' (login
      // screen showing), 'loading', 'unreachable' all keep the
      // browser source. For 'unauthenticated' the gate still hides
      // the editor so file ops can't fire; for the others the
      // browser source is the right anonymous default.
      setSource(setFileSourceKind('browser'));
    }
  }, [state.kind]);

  return <FileSourceContext.Provider value={source}>{children}</FileSourceContext.Provider>;
}

export function useFileSource(): FileSource {
  const value = useContext(FileSourceContext);
  if (!value) {
    // Defensive: callers below the provider never get null. If they do,
    // someone forgot to wrap in `FileSourceProvider`.
    return createBrowserFileSource();
  }
  return value;
}
