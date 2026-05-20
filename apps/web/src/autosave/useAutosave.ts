import { useEffect, useRef } from 'react';
import { ICommandService, type ICommandInfo, type IExecutionOptions } from '@univerjs/core';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { useCollab } from '../collab/collab-context';
import { clearAutosave, writeAutosave } from './store';

/**
 * Catches the "I closed the tab without saving" failure mode. Every
 * mutation marks the workbook dirty; an idle window (DEBOUNCE_MS) or
 * a regular tick (TICK_MS), whichever fires first, snapshots the
 * workbook into IndexedDB. A `pagehide` listener does one last
 * synchronous-ish save so unexpected tab-close is recoverable.
 *
 * Skipped in /r/<id> co-edit rooms — the room is already authoritative
 * server-side; restoring stale local state would actively confuse
 * the user when peers have moved on.
 *
 * The matching restore prompt lives in `<AutosaveRestoreBanner />`.
 */

const DEBOUNCE_MS = 5_000;
const TICK_MS = 30_000;

export function useAutosave(): void {
  const api = useUniverAPI();
  const workbook = useWorkbook();
  const collab = useCollab();
  const dirtyRef = useRef(false);
  const lastSaveRef = useRef(0);

  useEffect(() => {
    if (!api) return;
    if (collab.roomId) return; // covered by the room server

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const injector = (api as any)._injector as
      | { get: (t: unknown) => unknown }
      | undefined;
    if (!injector) return;
    const cmdSvc = injector.get(ICommandService) as {
      onMutationExecutedForCollab: (
        l: (info: ICommandInfo, options?: IExecutionOptions) => void,
      ) => { dispose: () => void };
    };

    let cancelled = false;
    let debounce: ReturnType<typeof setTimeout> | null = null;
    let tick: ReturnType<typeof setInterval> | null = null;

    const persist = async (reason: 'debounce' | 'tick' | 'pagehide') => {
      if (cancelled) return;
      if (!dirtyRef.current) return;
      const wb = api.getActiveWorkbook();
      if (!wb) return;
      const data = wb.save() as unknown as import('@univerjs/core').IWorkbookData;
      try {
        await writeAutosave({
          name: workbook.meta.name,
          sourceFormat: workbook.meta.sourceFormat ?? null,
          data,
          savedAt: Date.now(),
        });
        dirtyRef.current = false;
        lastSaveRef.current = Date.now();
        if (reason !== 'pagehide') console.debug('[autosave] saved', reason);
      } catch (err) {
        console.warn('[autosave] save failed', err);
      }
    };

    const scheduleDebounce = () => {
      if (debounce) clearTimeout(debounce);
      debounce = setTimeout(() => void persist('debounce'), DEBOUNCE_MS);
    };

    const subscription = cmdSvc.onMutationExecutedForCollab((info, options) => {
      // `fromCollab` mutations are remote replays — they don't dirty
      // the user's local copy in the autosave sense (the source of
      // truth is the room). Same as the collab bridge skip.
      if (options?.fromCollab) return;
      // Ignore the noisy selection / set-current-sheet mutations that
      // fire on every click — they don't change document state.
      const id = info?.id ?? '';
      if (id.startsWith('sheet.mutation.set-selections')) return;
      if (id === 'sheet.mutation.set-worksheet-active-operation') return;
      dirtyRef.current = true;
      scheduleDebounce();
    });

    tick = setInterval(() => void persist('tick'), TICK_MS);

    const onPageHide = () => {
      // pagehide fires on tab close + on bfcache stash. Try to flush
      // synchronously-ish — writeAutosave's IDB write isn't truly sync
      // but the browser keeps the page alive long enough for the put().
      void persist('pagehide');
    };
    window.addEventListener('pagehide', onPageHide);
    window.addEventListener('beforeunload', onPageHide);

    return () => {
      cancelled = true;
      subscription.dispose();
      if (debounce) clearTimeout(debounce);
      if (tick) clearInterval(tick);
      window.removeEventListener('pagehide', onPageHide);
      window.removeEventListener('beforeunload', onPageHide);
    };
    // workbook.meta is captured fresh on each save; not a dep to avoid
    // re-subscribing on every name edit.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api, collab.roomId]);
}

/** Imperatively drop the autosave slot — used after an explicit Save
 *  succeeds, since the on-disk file now supersedes the autosave. */
export async function discardAutosaveAfterExplicitSave(): Promise<void> {
  await clearAutosave();
}
