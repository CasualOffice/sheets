import { useEffect, useState } from 'react';
import type { IWorkbookData } from '@univerjs/core';
import { useWorkbook } from '../use-workbook';
import { useCollab } from '../collab/collab-context';
import { clearAutosave, readAutosave, type AutosaveRecord } from './store';

/**
 * Restore-prompt banner shown at app boot when an autosave record
 * exists from the last session. The user can:
 *
 *   - **Restore** — replace the empty workbook with the saved snapshot.
 *   - **Discard** — drop the saved record.
 *
 * Hidden while inside a /r/<id> co-edit room (the room is authoritative).
 *
 * Two gates keep the banner from showing for noise:
 *   - **Content gate**: the snapshot has to contain at least one cell
 *     with a value or formula. A workbook the user opened-then-closed
 *     without typing anything used to surface this banner; now it
 *     silently auto-discards.
 *   - **Age gate**: 24 h. Older snapshots are auto-discarded; if a user
 *     hasn't returned to a tab in a day the saved state is almost
 *     certainly stale.
 *
 * The banner self-dismisses after either action.
 */

const STALE_AFTER_MS = 24 * 60 * 60 * 1000; // 24 h

function countMeaningfulCells(data: IWorkbookData | null | undefined): number {
  if (!data?.sheets) return 0;
  let n = 0;
  for (const sheet of Object.values(data.sheets)) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const cellData = (sheet as any)?.cellData;
    if (!cellData || typeof cellData !== 'object') continue;
    for (const row of Object.values(cellData)) {
      if (!row || typeof row !== 'object') continue;
      for (const cell of Object.values(row as Record<string, unknown>)) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const c = cell as any;
        if (!c) continue;
        // Treat as a real cell only if it carries a value, formula, or
        // rich-text body — pure-style cells (a leftover border / font)
        // don't count as "the user typed something to restore".
        if (
          (c.v !== undefined && c.v !== null && c.v !== '') ||
          (typeof c.f === 'string' && c.f.length > 0) ||
          c.p
        ) {
          n += 1;
          if (n >= 3) return n; // early-out — we only need a low threshold
        }
      }
    }
  }
  return n;
}

export function AutosaveRestoreBanner() {
  const workbook = useWorkbook();
  const collab = useCollab();
  const [rec, setRec] = useState<AutosaveRecord | null>(null);
  const [dismissed, setDismissed] = useState(false);

  useEffect(() => {
    if (collab.roomId) return;
    let cancelled = false;
    void (async () => {
      const r = await readAutosave();
      if (cancelled) return;
      if (!r) {
        setRec(null);
        return;
      }
      // Age gate — drop stale snapshots silently.
      if (Date.now() - r.savedAt > STALE_AFTER_MS) {
        await clearAutosave();
        setRec(null);
        return;
      }
      // Content gate — drop "user opened a blank workbook and clicked
      // around" snapshots silently. The banner is only useful when
      // there's actual data to recover.
      if (countMeaningfulCells(r.data) === 0) {
        await clearAutosave();
        setRec(null);
        return;
      }
      setRec(r);
    })();
    return () => {
      cancelled = true;
    };
  }, [collab.roomId]);

  if (!rec || dismissed || collab.roomId) return null;

  const ago = formatAgo(Date.now() - rec.savedAt);

  return (
    <div className="autosave-banner" role="status" data-testid="autosave-banner">
      <span className="autosave-banner__text">
        Unsaved changes from <strong>{rec.name}</strong> ({ago}) — restore?
      </span>
      <button
        type="button"
        className="autosave-banner__btn autosave-banner__btn--primary"
        data-testid="autosave-restore"
        onClick={() => {
          workbook.replaceWorkbook(rec.data, rec.sourceFormat as Parameters<typeof workbook.replaceWorkbook>[1]);
          void clearAutosave();
          setDismissed(true);
        }}
      >
        Restore
      </button>
      <button
        type="button"
        className="autosave-banner__btn"
        data-testid="autosave-discard"
        onClick={() => {
          void clearAutosave();
          setDismissed(true);
        }}
      >
        Discard
      </button>
    </div>
  );
}

function formatAgo(ms: number): string {
  if (ms < 60_000) return 'moments ago';
  const mins = Math.round(ms / 60_000);
  if (mins < 60) return `${mins} min ago`;
  const hours = Math.round(mins / 60);
  if (hours < 24) return `${hours} hr ago`;
  const days = Math.round(hours / 24);
  return `${days} day${days === 1 ? '' : 's'} ago`;
}
