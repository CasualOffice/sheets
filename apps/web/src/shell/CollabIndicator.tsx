import { useCollab } from '../collab/collab-context';
import { Tooltip } from './Tooltip';

/**
 * Small status dot in the sheet-tabs strip showing the collab connection
 * state. Stays out of the way in single-user mode (compact "Solo" pill);
 * becomes a green live indicator with a peer count inside a room.
 *
 * Priority of override states (highest first):
 *   1. `replayFailures > 0` — amber. Remote mutations failed to apply
 *      locally, so our view is missing peer edits. "Refresh recommended."
 *   2. `syncHealth === 'diverged'` — amber. State vectors have
 *      disagreed for > 15 s. Refresh-recommended too.
 *   3. Transport status: live / connecting / offline / off.
 *
 * Both warning paths use the same "diverged" CSS class so the visual
 * stays consistent (one shade of amber for "you should refresh").
 */
export function CollabIndicator() {
  const { status, roomId, syncHealth, peerCount, queuedLocal, replayFailures } = useCollab();
  const failed = status === 'live' && replayFailures > 0;
  const diverged = !failed && status === 'live' && syncHealth === 'diverged';
  const effectiveStatus = failed || diverged ? 'diverged' : status;

  // Visible text on the pill. Kept short — the tooltip carries the
  // full context. "Live · 2" reads as "live, with two others"; "Solo"
  // makes the single-user case unambiguous.
  let text: string;
  if (failed) text = `${replayFailures} not synced`;
  else if (diverged) text = 'Out of sync';
  else if (status === 'live') text = peerCount > 0 ? `Live · ${peerCount}` : 'Live';
  else if (status === 'connecting') text = '…';
  else if (status === 'offline')
    text = queuedLocal > 0 ? `Reconnecting · ${queuedLocal}` : 'Reconnecting…';
  else text = 'Solo';

  // Tooltip — the long-form version of whatever the pill says, plus
  // the roomId for the share-link case.
  let baseLabel: string;
  if (failed) {
    baseLabel = `${replayFailures} ${replayFailures === 1 ? 'edit from a peer' : 'edits from peers'} couldn't be applied to your view — refresh to resync`;
  } else if (diverged) {
    baseLabel = 'Out of sync with peers — refresh usually recovers';
  } else if (status === 'live') {
    baseLabel =
      peerCount > 0
        ? `Live — co-editing with ${peerCount} ${peerCount === 1 ? 'other peer' : 'other peers'}`
        : 'Live — co-editing on, no one else here yet';
  } else if (status === 'connecting') {
    baseLabel = 'Connecting to room…';
  } else if (status === 'offline') {
    baseLabel =
      queuedLocal > 0
        ? `Reconnecting — ${queuedLocal} of your ${queuedLocal === 1 ? 'change' : 'changes'} queued locally; they'll sync when the connection is back`
        : 'Reconnecting — your edits will sync when the connection is back';
  } else {
    baseLabel = 'Single-user mode';
  }
  const label = roomId ? `${baseLabel} (room ${roomId})` : baseLabel;

  return (
    <Tooltip label={label} side="top">
      <span
        className={`collab-indicator collab-indicator--${effectiveStatus}`}
        data-testid="collab-indicator"
        data-collab-status={effectiveStatus}
        data-sync-health={syncHealth}
        data-peer-count={peerCount}
        data-queued-local={queuedLocal}
        data-replay-failures={replayFailures}
        role="status"
        aria-label={label}
      >
        <span className="collab-indicator__dot" aria-hidden="true" />
        <span className="collab-indicator__text">{text}</span>
      </span>
    </Tooltip>
  );
}
