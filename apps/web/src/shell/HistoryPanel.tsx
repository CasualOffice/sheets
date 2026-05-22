import { useEffect, useState } from 'react';
import * as Y from 'yjs';
import { useCollab } from '../collab/collab-context';
import { usePresence } from '../collab/presence-context';
import { Icon } from './Icon';

/**
 * Live co-edit history side panel. Reads the bridge's Yjs op-log
 * (`doc.getArray('ops')`) and renders entries as "who changed what
 * when". Session-scoped: when the tab closes the panel forgets, by
 * design — we don't want a side channel for persisted change history
 * that survives past the live room.
 *
 * v1 is read-only. Revert is a planned follow-up; the data shape on
 * disk (clientId + mutation id + params + timestamp) is sufficient to
 * issue an inverse mutation, but the inverse-mutation factory work
 * (especially for structural ops like insert-row that need OT) is
 * out of scope for this panel pass.
 */

type LogRecord = {
  c: string;
  t: number;
  id?: string;
  kind?: 'snapshot';
  p?: unknown;
};

export function HistoryPanel() {
  const { doc, roomId } = useCollab();
  const { me, peers } = usePresence();
  const [entries, setEntries] = useState<LogRecord[]>([]);

  // Subscribe to the op-log array. Re-snapshot on every change; the
  // array is small enough (capped by compaction) that a full re-read
  // per change is fine, and avoids the bookkeeping of incremental
  // diffs against React state.
  useEffect(() => {
    if (!doc) {
      setEntries([]);
      return;
    }
    const log = doc.getArray<LogRecord>('ops');
    const snapshot = () => setEntries(log.toArray());
    snapshot();
    const observer = (_ev: Y.YArrayEvent<LogRecord>) => snapshot();
    log.observe(observer);
    return () => log.unobserve(observer);
  }, [doc]);

  if (!roomId) {
    return (
      <aside className="side-panel" data-testid="history-panel" aria-label="Session history">
        <header className="side-panel__header">
          <Icon name="history" size="sm" />
          <h2 className="side-panel__title">History</h2>
        </header>
        <div className="side-panel__empty">
          History is only available inside a co-edit room. Share the workbook to start a session.
        </div>
      </aside>
    );
  }

  // Build a clientId → name lookup. The local doc's clientId is on
  // the Yjs doc itself; peer names come from awareness.
  const clientNames = new Map<string, string>();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const myCid = doc ? String((doc as any).clientID) : '';
  if (me) clientNames.set(myCid, `${me.name} (you)`);
  for (const p of peers) clientNames.set(String(p.clientId), p.name);

  // Render newest first — humans scan downward from the top.
  const sorted = [...entries].sort((a, b) => b.t - a.t);

  return (
    <aside className="side-panel" data-testid="history-panel" aria-label="Session history">
      <header className="side-panel__header">
        <Icon name="history" size="sm" />
        <h2 className="side-panel__title">History</h2>
        <span className="side-panel__count" data-testid="history-count">
          {entries.length}
        </span>
      </header>
      {sorted.length === 0 ? (
        <div className="side-panel__empty">
          No changes yet in this session. Type into a cell to start recording.
        </div>
      ) : (
        <ol className="history-list" role="list">
          {sorted.map((rec, i) => (
            <li key={i} className="history-list__row" data-testid="history-row">
              <div className="history-list__who" style={{ color: peerColor(rec.c, me?.name, peers) }}>
                {clientNames.get(rec.c) ?? `client ${rec.c.slice(0, 6)}`}
              </div>
              <div className="history-list__what">{describe(rec)}</div>
              <time className="history-list__when" dateTime={new Date(rec.t).toISOString()}>
                {formatTime(rec.t)}
              </time>
            </li>
          ))}
        </ol>
      )}
    </aside>
  );
}

/**
 * Map a log entry to a one-line user-facing description.
 * Mutation params are typed implicitly by the mutation id; this
 * switch keeps the strings honest about what each id means without
 * pulling the whole Univer mutation-param type tree in.
 */
function describe(rec: LogRecord): string {
  if (rec.kind === 'snapshot') return 'Compacted history — earlier changes folded into a snapshot';
  const id = rec.id ?? '';
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const p = (rec.p ?? {}) as any;
  switch (id) {
    case 'sheet.mutation.set-range-values': {
      const cellCount = countCells(p?.cellValue);
      return cellCount === 1 ? 'Edited 1 cell' : `Edited ${cellCount} cells`;
    }
    case 'sheet.mutation.set-style':
      return 'Changed cell style';
    case 'sheet.mutation.insert-row':
      return `Inserted ${p?.rowCount ?? ''} row(s)`.trim();
    case 'sheet.mutation.insert-col':
      return `Inserted ${p?.colCount ?? ''} column(s)`.trim();
    case 'sheet.mutation.remove-row':
      return 'Deleted row(s)';
    case 'sheet.mutation.remove-col':
      return 'Deleted column(s)';
    case 'sheet.mutation.move-rows':
      return 'Moved row(s)';
    case 'sheet.mutation.move-cols':
      return 'Moved column(s)';
    case 'sheet.mutation.set-row-hidden':
      return 'Hid row(s)';
    case 'sheet.mutation.set-row-visible':
      return 'Showed row(s)';
    case 'sheet.mutation.set-col-hidden':
      return 'Hid column(s)';
    case 'sheet.mutation.set-col-visible':
      return 'Showed column(s)';
    case 'sheet.mutation.set-worksheet-row-height':
      return 'Resized row(s)';
    case 'sheet.mutation.set-worksheet-col-width':
      return 'Resized column(s)';
    case 'sheet.mutation.add-worksheet-merge':
      return 'Merged cells';
    case 'sheet.mutation.remove-worksheet-merge':
      return 'Unmerged cells';
    case 'sheet.mutation.insert-sheet':
      return `Added sheet "${p?.sheet?.name ?? 'Sheet'}"`;
    case 'sheet.mutation.remove-sheet':
      return 'Removed a sheet';
    case 'sheet.mutation.set-worksheet-name':
      return `Renamed sheet to "${p?.name ?? ''}"`;
    case 'sheet.mutation.set-worksheet-order':
      return 'Reordered sheets';
    case 'sheet.mutation.set-worksheet-hidden':
      return p?.hidden ? 'Hid a sheet' : 'Showed a sheet';
    case 'sheet.mutation.set-frozen':
      return 'Changed freeze panes';
    case 'sheet.mutation.add-hyper-link':
      return 'Added a hyperlink';
    case 'sheet.mutation.remove-hyper-link':
      return 'Removed a hyperlink';
    case 'sheet.mutation.update-hyper-link':
      return 'Updated a hyperlink';
    default:
      return id.replace(/^sheet\.mutation\./, '') || 'Change';
  }
}

function countCells(cellValue: unknown): number {
  if (!cellValue || typeof cellValue !== 'object') return 0;
  let n = 0;
  for (const row of Object.values(cellValue as Record<string, unknown>)) {
    if (row && typeof row === 'object') n += Object.keys(row as Record<string, unknown>).length;
  }
  return n;
}

function formatTime(ts: number): string {
  const now = Date.now();
  const delta = now - ts;
  if (delta < 5_000) return 'just now';
  if (delta < 60_000) return `${Math.floor(delta / 1000)}s ago`;
  if (delta < 3_600_000) return `${Math.floor(delta / 60_000)}m ago`;
  // For older entries, show absolute clock time — session-scoped anyway
  // so a full datetime would be noise.
  const d = new Date(ts);
  return d.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
}

function peerColor(
  clientId: string,
  myName: string | undefined,
  peers: Array<{ clientId: number; color: string }>,
): string | undefined {
  for (const p of peers) {
    if (String(p.clientId) === clientId) return p.color;
  }
  // For the local user we don't have a stored color on context here;
  // fall back to the default text colour (myName is unused after the
  // earlier rewrite — kept in the signature for future "highlight my
  // edits" UX).
  void myName;
  return undefined;
}
