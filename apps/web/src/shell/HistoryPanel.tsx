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

import { useEffect, useState } from 'react';
import type * as Y from 'yjs';
import { useCollab } from '../collab/collab-context';
import { usePresence } from '../collab/presence-context';
import { useUniverAPI } from '../use-univer';
import { useLocalHistory } from './local-history';
import { Icon } from './Icon';

/**
 * Session history side panel. Drives off two different sources depending
 * on mode:
 *
 *   - **Co-edit room**: reads the bridge's Yjs op-log
 *     (`doc.getArray('ops')`) — shared across peers, scoped to the
 *     live room.
 *   - **Solo session**: reads an in-memory ring fed by the local
 *     ICommandService (see `useLocalHistory`). Same record shape;
 *     same revert path; entries don't survive a refresh.
 *
 * Revert works for both sources: cell-value (set-range-values) entries
 * carry undo params captured before the redo ran. Structural ops
 * (insert-row, move-range, …) need their own inverse factories and
 * stay non-revertable for v1; they still render in the timeline.
 */

type LogRecord = {
  c: string;
  t: number;
  id?: string;
  kind?: 'snapshot';
  p?: unknown;
  /** Captured undo params (set for REVERTABLE_MUTATIONS in
   *  bridge.ts / local-history.ts). When present, Revert is enabled. */
  u?: unknown;
};

export function HistoryPanel() {
  const { doc, roomId } = useCollab();
  const { me, peers } = usePresence();
  const api = useUniverAPI();
  const [collabEntries, setCollabEntries] = useState<LogRecord[]>([]);
  const [reverting, setReverting] = useState<number | null>(null);

  // Collab op-log subscription. Inactive when there is no Yjs doc.
  useEffect(() => {
    if (!doc) {
      setCollabEntries([]);
      return;
    }
    const log = doc.getArray<LogRecord>('ops');
    const snapshot = () => setCollabEntries(log.toArray());
    snapshot();
    const observer = (_ev: Y.YArrayEvent<LogRecord>) => snapshot();
    log.observe(observer);
    return () => log.unobserve(observer);
  }, [doc]);

  // Local mutation feed — only active in solo sessions. The hook
  // returns an empty array when api is null, so it's safe to call
  // unconditionally and pick the right source after.
  const localEntries = useLocalHistory(roomId ? null : api);
  const entries: LogRecord[] = roomId ? collabEntries : localEntries;

  // Build a clientId → name lookup. The local doc's clientId is on
  // the Yjs doc itself; peer names come from awareness. Solo mode has
  // only the `me` slot keyed by the placeholder client id used by
  // useLocalHistory.
  const clientNames = new Map<string, string>();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const myCid = doc ? String((doc as any).clientID) : 'me';
  clientNames.set(myCid, me ? `${me.name} (you)` : 'You');
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
          {sorted.map((rec, i) => {
            const canRevert = rec.u !== undefined && rec.id !== undefined && !!api;
            const onRevert = async () => {
              if (!canRevert) return;
              setReverting(i);
              try {
                // Execute the captured undo params LOCALLY (no
                // fromCollab) so the bridge captures it as a new
                // mutation and propagates it to peers — the revert
                // becomes a normal edit in everyone's history.
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const apiAny = api as any;
                await apiAny.executeCommand?.(rec.id, rec.u);
              } catch (err) {
                console.warn('[history] revert failed for', rec.id, err);
              } finally {
                setReverting(null);
              }
            };
            return (
              <li key={i} className="history-list__row" data-testid="history-row">
                <div
                  className="history-list__who"
                  style={{ color: peerColor(rec.c, me?.name, peers) }}
                >
                  {clientNames.get(rec.c) ?? `client ${rec.c.slice(0, 6)}`}
                </div>
                <time
                  className="history-list__when"
                  dateTime={new Date(rec.t).toISOString()}
                >
                  {formatTime(rec.t)}
                </time>
                <div className="history-list__what">{describe(rec)}</div>
                <button
                  type="button"
                  className="history-list__revert"
                  data-testid="history-revert"
                  disabled={!canRevert || reverting !== null}
                  title={
                    canRevert
                      ? 'Revert this change — applies as a new edit, also visible to peers.'
                      : 'Revert is only available for cell-value changes captured after this session opened.'
                  }
                  onClick={() => void onRevert()}
                >
                  {reverting === i ? 'Reverting…' : 'Revert'}
                </button>
              </li>
            );
          })}
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
