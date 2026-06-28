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

import { useCallback, useEffect, useState } from 'react';
import {
  getFolderState,
  pinFolder,
  reconnectFolder,
  unpinFolder,
  type FolderState,
} from './pinned-folder';

/**
 * Reactive folder-pin state for the home screen and the File menu.
 *
 * `state.kind` drives which control we render:
 *   - 'unsupported' → don't show pinning UI at all
 *   - 'none'        → "Pin a folder" button
 *   - 'granted'     → "Save to <name>" badge + "Unpin"
 *   - 'prompt'      → "Reconnect <name>" button (permission lapsed)
 *   - 'denied'      → "Re-pin <name>" (user denied; offer to pick again)
 */

export function usePinnedFolder() {
  const [state, setState] = useState<FolderState>({ kind: 'none' });

  const refresh = useCallback(async () => {
    const next = await getFolderState();
    setState(next);
  }, []);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  const pin = useCallback(async () => {
    const rec = await pinFolder();
    if (rec) setState({ kind: 'granted', record: rec });
    return rec;
  }, []);

  const reconnect = useCallback(async () => {
    if (state.kind !== 'prompt' && state.kind !== 'denied') return false;
    const ok = await reconnectFolder(state.record);
    if (ok) setState({ kind: 'granted', record: state.record });
    return ok;
  }, [state]);

  const unpin = useCallback(async () => {
    await unpinFolder();
    setState({ kind: 'none' });
  }, []);

  return { state, refresh, pin, reconnect, unpin };
}
