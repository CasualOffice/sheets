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

import { useEffect } from 'react';
import {
  ICommandService,
  type ICommandInfo,
  type IExecutionOptions,
  type IWorkbookData,
} from '@univerjs/core';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { useCollab } from '../collab/collab-context';

/**
 * Desktop crash-recovery — the native-app analogue of {@link useAutosave}.
 *
 * Where autosave persists to IndexedDB (per-browser, gone if the profile dir
 * is wiped), this writes the unsaved-changes snapshot to a disk sidecar the
 * Rust shell owns, so the launcher can offer "Recover unsaved changes?" across
 * windows and process restarts. It shares autosave's change hook, debounce and
 * noisy-mutation filter so the two agree on what "an edit" is.
 *
 * The snapshot is the Univer `IWorkbookData` (its native, lossless form)
 * serialized as JSON bytes — NOT a re-encoded `.xlsx`, which would round-trip
 * through ExcelJS and drop Univer-specific state. Restore feeds it straight
 * back through `replaceWorkbook`, exactly like the autosave restore banner.
 *
 * Desktop-only: a no-op on web, in `/r/<id>` co-edit rooms (the room server is
 * authoritative), and while untitled (the bridge has no path to key on — the
 * bridge methods themselves no-op in that case).
 */

export interface DesktopRecoveryRecord {
  name: string;
  sourceFormat: string | null;
  data: IWorkbookData;
  savedAt: number;
}

function deskBridge() {
  return typeof window !== 'undefined' ? window.__deskApp__ : undefined;
}

const encoder = new TextEncoder();
const decoder = new TextDecoder('utf-8');

// Whether there are edits worth snapshotting. Module-level (not the writer
// hook's local ref) so clearDesktopRecovery — invoked from the save chokepoint
// OUTSIDE the hook — can reset it: a debounced snapshot still pending when a
// clean Save fires must NOT resurrect the just-cleared sidecar.
let recoveryDirty = false;
// Bumped on every edit AND every clear. persist() captures it before its async
// write and only marks clean afterwards if it's unchanged — so an edit (or a
// Save's clear) that lands mid-write isn't silently dropped from the next snap.
let recoveryGen = 0;

export async function writeDesktopRecovery(record: DesktopRecoveryRecord): Promise<void> {
  const b = deskBridge();
  if (!b?.writeRecovery) return;
  // TextEncoder.encode returns a Uint8Array backed by a fresh, exactly-sized
  // ArrayBuffer (byteOffset 0), so handing `.buffer` straight through is safe.
  const u8 = encoder.encode(JSON.stringify(record));
  await b.writeRecovery(u8.buffer as ArrayBuffer);
}

export async function readDesktopRecovery(): Promise<DesktopRecoveryRecord | null> {
  const b = deskBridge();
  if (!b?.readRecovery) return null;
  const raw = await b.readRecovery();
  if (!raw || raw.byteLength === 0) return null;
  try {
    return JSON.parse(decoder.decode(raw)) as DesktopRecoveryRecord;
  } catch {
    return null; // corrupt sidecar — treat as no recovery
  }
}

export async function clearDesktopRecovery(): Promise<void> {
  // Reset synchronously, before the await, so a debounced snapshot that fires
  // while the clear IPC is in flight sees a clean state and skips its write.
  recoveryDirty = false;
  recoveryGen += 1;
  const b = deskBridge();
  if (!b?.clearRecovery) return;
  try {
    await b.clearRecovery();
  } catch {
    /* best-effort — clearing the sidecar must never surface as an error */
  }
}

const DEBOUNCE_MS = 5_000;
const TICK_MS = 30_000;

/**
 * Writer side: subscribe to edits, debounce, snapshot the workbook to the
 * host sidecar. Mounted as an effect-only driver alongside the autosave one.
 */
export function useDesktopRecoveryWriter(): void {
  const api = useUniverAPI();
  const workbook = useWorkbook();
  const collab = useCollab();

  useEffect(() => {
    if (!api) return;
    if (!deskBridge()?.isDesktop) return; // web — no native sidecar
    if (collab.roomId) return; // room server owns persistence

    // Reach the command service via the facade's private `_injector` — the
    // same back door useAutosave / EditTracker use, so all three stay aligned
    // on what counts as an edit (CLAUDE.md: onMutationExecutedForCollab is the
    // only sanctioned change hook).
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const injector = (api as any)._injector as { get: (t: unknown) => unknown } | undefined;
    if (!injector) return;
    const cmdSvc = injector.get(ICommandService) as {
      onMutationExecutedForCollab: (
        l: (info: ICommandInfo, options?: IExecutionOptions) => void,
      ) => { dispose: () => void };
    };

    let cancelled = false;
    let debounce: ReturnType<typeof setTimeout> | null = null;
    let tick: ReturnType<typeof setInterval> | null = null;

    const persist = async () => {
      if (cancelled || !recoveryDirty) return;
      const wb = api.getActiveWorkbook();
      if (!wb) return;
      const gen = recoveryGen;
      const data = wb.save() as unknown as IWorkbookData;
      try {
        await writeDesktopRecovery({
          name: workbook.meta.name,
          sourceFormat: workbook.meta.sourceFormat ?? null,
          data,
          savedAt: Date.now(),
        });
        // Only mark clean if nothing changed (an edit OR a Save's clear) while
        // the snapshot was being written — otherwise we'd drop the newer edit.
        if (recoveryGen === gen) recoveryDirty = false;
      } catch (err) {
        // Best-effort — a recovery snapshot must never disrupt editing.
        console.debug('[deskApp] recovery snapshot failed', err);
      }
    };

    const scheduleDebounce = () => {
      if (debounce) clearTimeout(debounce);
      debounce = setTimeout(() => void persist(), DEBOUNCE_MS);
    };

    const sub = cmdSvc.onMutationExecutedForCollab((info, options) => {
      if (options?.fromCollab) return; // remote replays don't count
      const id = info?.id ?? '';
      // Same noisy-mutation filter as useAutosave.
      if (id.startsWith('sheet.mutation.set-selections')) return;
      if (id === 'sheet.mutation.set-worksheet-active-operation') return;
      recoveryDirty = true;
      recoveryGen += 1;
      scheduleDebounce();
    });

    tick = setInterval(() => void persist(), TICK_MS);

    return () => {
      cancelled = true;
      sub.dispose();
      if (debounce) clearTimeout(debounce);
      if (tick) clearInterval(tick);
    };
    // workbook.meta is read fresh on each persist; intentionally not a dep so
    // we don't re-subscribe on every name edit. Mirrors useAutosave.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api, collab.roomId]);
}
