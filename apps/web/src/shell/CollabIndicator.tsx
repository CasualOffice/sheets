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

import { useEffect, useRef, useState } from 'react';
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
 *
 * When there are dead-letter entries (mutations the bridge gave up
 * on after the retry budget), the pill becomes click-to-expand: a
 * popover lists the latest 5 entries with mutation id, classification,
 * truncated error, and age. Lets the user (and us, in production)
 * self-diagnose what's actually failing instead of just seeing a
 * count. See docs/PRODUCTION_PIPELINE.md → Stream A2.
 */
export function CollabIndicator() {
  const { status, roomId, syncHealth, peerCount, queuedLocal, replayFailures, replayDeadLetter } =
    useCollab();
  const failed = status === 'live' && replayFailures > 0;
  const diverged = !failed && status === 'live' && syncHealth === 'diverged';
  const effectiveStatus = failed || diverged ? 'diverged' : status;
  const hasDetail = replayDeadLetter.length > 0;
  const [expanded, setExpanded] = useState(false);

  // Auto-close the popover if the failure state clears (e.g. user
  // refreshed and the bridge restarted with a fresh empty buffer).
  useEffect(() => {
    if (!hasDetail) setExpanded(false);
  }, [hasDetail]);

  // Close on outside click / Escape — popover hygiene.
  const containerRef = useRef<HTMLSpanElement | null>(null);
  useEffect(() => {
    if (!expanded) return;
    const onDoc = (e: MouseEvent) => {
      if (!containerRef.current) return;
      if (containerRef.current.contains(e.target as Node)) return;
      setExpanded(false);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setExpanded(false);
    };
    document.addEventListener('mousedown', onDoc);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('mousedown', onDoc);
      document.removeEventListener('keydown', onKey);
    };
  }, [expanded]);

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
    const verb = hasDetail ? 'click for detail' : 'refresh to resync';
    baseLabel = `${replayFailures} ${replayFailures === 1 ? 'edit from a peer' : 'edits from peers'} couldn't be applied to your view — ${verb}`;
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

  const pill = (
    <Tooltip label={label} side="top">
      {hasDetail ? (
        <button
          type="button"
          className={`collab-indicator collab-indicator--${effectiveStatus} collab-indicator--clickable`}
          data-testid="collab-indicator"
          data-collab-status={effectiveStatus}
          data-sync-health={syncHealth}
          data-peer-count={peerCount}
          data-queued-local={queuedLocal}
          data-replay-failures={replayFailures}
          data-dead-letter-count={replayDeadLetter.length}
          aria-label={label}
          aria-expanded={expanded}
          aria-haspopup="dialog"
          onClick={() => setExpanded((v) => !v)}
        >
          <span className="collab-indicator__dot" aria-hidden="true" />
          <span className="collab-indicator__text">{text}</span>
        </button>
      ) : (
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
      )}
    </Tooltip>
  );

  return (
    <span className="collab-indicator__wrap" ref={containerRef} data-testid="collab-indicator-wrap">
      {pill}
      {expanded && hasDetail && <ReplayFailurePopover entries={replayDeadLetter} />}
    </span>
  );
}

interface PopoverProps {
  entries: readonly import('@casualoffice/sheets/collab').ReplayFailureRecord[];
}

/**
 * Click-to-expand detail panel for the dead-letter ring buffer.
 * Renders the last 5 entries in reverse-chronological order with:
 *   - mutation id (e.g. "sheet.mutation.set-range-values")
 *   - classification chip (transient / permanent)
 *   - truncated error message
 *   - age relative to now
 *
 * Deliberately read-only — there's no "retry" button. Permanent
 * failures can't be retried meaningfully (same error), and transient
 * failures already exhausted their retry budget by the time they
 * land here. The recovery action is "refresh", surfaced in the
 * tooltip.
 */
function ReplayFailurePopover({ entries }: PopoverProps) {
  const now = Date.now();
  const latest = entries.slice(-5).reverse();
  return (
    <div
      className="collab-indicator__popover"
      role="dialog"
      aria-label="Replay failures"
      data-testid="replay-failure-popover"
    >
      <header className="collab-indicator__popover-header">
        <strong>Replay failures</strong>
        <span className="collab-indicator__popover-hint">
          Showing latest {latest.length} of {entries.length}
        </span>
      </header>
      <ol className="collab-indicator__popover-list">
        {latest.map((entry, idx) => (
          <li
            key={`${entry.firstFailedAt}-${idx}`}
            className="collab-indicator__popover-item"
            data-classification={entry.classification}
          >
            <div className="collab-indicator__popover-row">
              <code className="collab-indicator__popover-id">{shortenMutationId(entry.id)}</code>
              <span
                className={`collab-indicator__popover-chip collab-indicator__popover-chip--${entry.classification}`}
              >
                {entry.classification}
              </span>
              <span className="collab-indicator__popover-age">
                {formatAge(now - entry.lastFailedAt)}
              </span>
            </div>
            <div className="collab-indicator__popover-err" title={entry.lastError}>
              {truncate(entry.lastError, 140)}
            </div>
          </li>
        ))}
      </ol>
      <footer className="collab-indicator__popover-footer">
        Refresh usually recovers. If failures persist, copy a sample and file an issue.
      </footer>
    </div>
  );
}

function shortenMutationId(id: string): string {
  // Trim the verbose "sheet.mutation." prefix that ~all entries
  // share — leaves the distinguishing suffix readable in the
  // narrow popover column.
  return id.replace(/^sheet\.mutation\./, '');
}

function truncate(s: string, max: number): string {
  if (s.length <= max) return s;
  return `${s.slice(0, max - 1)}…`;
}

function formatAge(ms: number): string {
  if (ms < 0) return 'just now';
  const s = Math.round(ms / 1000);
  if (s < 5) return 'just now';
  if (s < 60) return `${s}s ago`;
  const m = Math.round(s / 60);
  if (m < 60) return `${m}m ago`;
  const h = Math.round(m / 60);
  return `${h}h ago`;
}
