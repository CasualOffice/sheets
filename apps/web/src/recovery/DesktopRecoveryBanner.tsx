import { useEffect, useState } from 'react';
import { useWorkbook } from '../use-workbook';
import { useCollab } from '../collab/collab-context';
import { useToast } from '../shell/toast/toast-context';
import {
  clearDesktopRecovery,
  readDesktopRecovery,
  type DesktopRecoveryRecord,
} from './desktop-recovery';

/**
 * Desktop crash-recovery restore prompt — the native-app sibling of
 * `<AutosaveRestoreBanner />`, reusing its `.autosave-banner` styling.
 *
 * On boot it reads the host recovery sidecar for the bound file. A sidecar
 * only survives when the previous session was killed mid-edit (a clean Save
 * clears it), so its mere existence means there are unsaved changes to offer —
 * no age / content gate needed. Restore swaps the snapshot in via
 * `replaceWorkbook` and clears the sidecar; Discard just clears it.
 *
 * Hidden inside a `/r/<id>` co-edit room (the room is authoritative) and on
 * web (the bridge methods return null, so `rec` stays null).
 */
export function DesktopRecoveryBanner() {
  const workbook = useWorkbook();
  const collab = useCollab();
  const toast = useToast();
  const [rec, setRec] = useState<DesktopRecoveryRecord | null>(null);
  const [dismissed, setDismissed] = useState(false);

  useEffect(() => {
    if (collab.roomId) return;
    let cancelled = false;
    void (async () => {
      const r = await readDesktopRecovery();
      if (!cancelled) setRec(r);
    })();
    return () => {
      cancelled = true;
    };
  }, [collab.roomId]);

  if (!rec || dismissed || collab.roomId) return null;

  const doRestore = () => {
    workbook.replaceWorkbook(
      rec.data,
      rec.sourceFormat as Parameters<typeof workbook.replaceWorkbook>[1],
    );
    toast.success(`Restored unsaved changes to ${rec.name}`);
    // Intentionally do NOT clear the sidecar here. The restored content is
    // unsaved (it differs from disk); replaceWorkbook re-dirties the workbook so
    // the writer re-snapshots it, and a clean Save clears it. Clearing now would
    // leave an immediate re-crash with nothing to recover.
    setDismissed(true);
  };

  return (
    <div className="autosave-banner" role="status" data-testid="desktop-recovery-banner">
      <span className="autosave-banner__text">
        <strong>{rec.name}</strong> had unsaved changes from a previous session — restore?
      </span>
      <button
        type="button"
        className="autosave-banner__btn autosave-banner__btn--primary"
        data-testid="desktop-recovery-restore"
        onClick={doRestore}
      >
        Restore
      </button>
      <button
        type="button"
        className="autosave-banner__btn"
        data-testid="desktop-recovery-discard"
        onClick={() => {
          void clearDesktopRecovery();
          setDismissed(true);
        }}
      >
        Discard
      </button>
    </div>
  );
}
