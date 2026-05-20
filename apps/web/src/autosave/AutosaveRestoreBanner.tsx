import { useEffect, useState } from 'react';
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
 * The banner self-dismisses after either action and remembers the
 * decision so it doesn't re-appear if the same tab reloads without
 * an intervening mutation.
 */

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
      if (!cancelled) setRec(r);
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
