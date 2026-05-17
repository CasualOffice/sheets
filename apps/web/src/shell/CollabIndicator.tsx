import { useCollab } from '../collab/collab-context';
import { Tooltip } from './Tooltip';

const LABELS: Record<string, string> = {
  off: 'Single-user mode',
  connecting: 'Connecting to room…',
  live: 'Live — co-editing on',
  offline: 'Disconnected — trying to reconnect',
};

/**
 * Small status dot in the sheet-tabs strip showing the collab connection
 * state. Stays out of the way in single-user mode (compact "single-user"
 * pill); becomes a green live indicator inside a room.
 *
 * Clicking nothing yet — pure status surface. A future iteration can make
 * the dot a popover with the share URL + connected-peer count.
 */
export function CollabIndicator() {
  const { status, roomId } = useCollab();
  const label = roomId ? `${LABELS[status]} (room ${roomId})` : LABELS[status];

  return (
    <Tooltip label={label} side="top">
      <span
        className={`collab-indicator collab-indicator--${status}`}
        data-testid="collab-indicator"
        data-collab-status={status}
        role="status"
        aria-label={label}
      >
        <span className="collab-indicator__dot" aria-hidden="true" />
        <span className="collab-indicator__text">
          {status === 'live' ? 'Live' : status === 'connecting' ? '…' : status === 'offline' ? 'Offline' : 'Solo'}
        </span>
      </span>
    </Tooltip>
  );
}
