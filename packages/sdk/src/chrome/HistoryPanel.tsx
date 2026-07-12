/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License").
 */

/**
 * History / Activity side panel — a self-contained, in-memory activity feed
 * of user-facing edits in the current session.
 *
 * Ported from the standalone app's HistoryPanel, but DECOUPLED from every
 * app-only coupling: the app version drove off two host-owned sources (a
 * shared Yjs op-log and IndexedDB-backed snapshot version history via
 * `useCollab`/`usePresence`/`local-history`). None of that exists in the SDK,
 * and persistent version history stays HOST-OWNED. So instead this panel
 * subscribes to the live `ICommandService.onCommandExecuted` stream, keeps a
 * bounded in-memory ring (last 100) of user-facing command executions, maps
 * each command id to a friendly label, and renders them newest-first.
 *
 * Ordering is derived from a monotonically incrementing sequence counter
 * (NOT wall-clock time) so nothing calls `Date.now()`/`new Date()` at module
 * load; a per-entry relative-age string is computed lazily at render time from
 * a capture timestamp taken inside the (post-mount) event handler.
 *
 * The app-only `useUI().toggleHistoryPanel` is replaced by the panel host's
 * `onClose`; the Univer facade + services are reached through the `api` prop's
 * `_injector`, matching StatusBar / FormulaBar precedent.
 */
import { useEffect, useMemo, useRef, useState, type CSSProperties } from 'react';
import { CommandType, ICommandService, type ICommandInfo } from '@univerjs/core';

import type { PanelComponentProps } from './extensions';
import { Icon } from './Icon';
import { PanelHeader } from './panel-shell';

const RING_CAP = 100;

interface ActivityEntry {
  /** Monotonic sequence — sole source of ordering (no wall-clock at module load). */
  seq: number;
  /** Command id that fired (e.g. `sheet.command.set-range-values`). */
  id: string;
  /** Friendly one-line label. */
  label: string;
  /** Icon glyph for the row. */
  icon: string;
  /** Capture time (ms). Taken inside the handler, post-mount — safe. */
  t: number;
}

/**
 * Map a user-facing COMMAND id to a friendly label + icon. Commands (not
 * mutations) are the user-intent layer, so one command == one activity row —
 * no de-duping a single paste into a dozen mutation rows. Anything not mapped
 * still records with a de-namespaced fallback label so the feed stays honest.
 */
const COMMAND_LABELS: Record<string, { label: string; icon: string }> = {
  'sheet.command.set-range-values': { label: 'Edited cells', icon: 'edit' },
  'sheet.command.set-range-bold': { label: 'Toggled bold', icon: 'format_bold' },
  'sheet.command.set-range-italic': { label: 'Toggled italic', icon: 'format_italic' },
  'sheet.command.set-range-underline': { label: 'Toggled underline', icon: 'format_underlined' },
  'sheet.command.set-range-stroke-through': {
    label: 'Toggled strikethrough',
    icon: 'strikethrough_s',
  },
  'sheet.command.set-range-font-family': { label: 'Changed font', icon: 'font_download' },
  'sheet.command.set-range-font-size': { label: 'Changed font size', icon: 'format_size' },
  'sheet.command.set-range-text-color': { label: 'Changed text color', icon: 'format_color_text' },
  'sheet.command.set-background-color': { label: 'Changed fill color', icon: 'format_color_fill' },
  'sheet.command.set-range-text-align': { label: 'Changed alignment', icon: 'format_align_left' },
  'sheet.command.set-range-vertical-align': {
    label: 'Changed alignment',
    icon: 'vertical_align_center',
  },
  'sheet.command.set-border-command': { label: 'Changed borders', icon: 'border_all' },
  'sheet.command.clear-selection-content': { label: 'Cleared cell contents', icon: 'backspace' },
  'sheet.command.clear-selection-format': { label: 'Cleared formatting', icon: 'format_clear' },
  'sheet.command.clear-selection-all': { label: 'Cleared cells', icon: 'clear_all' },
  'sheet.command.insert-row': { label: 'Inserted row(s)', icon: 'add_row_above' },
  'sheet.command.insert-row-before': { label: 'Inserted row above', icon: 'add_row_above' },
  'sheet.command.insert-row-after': { label: 'Inserted row below', icon: 'add_row_below' },
  'sheet.command.insert-col': { label: 'Inserted column(s)', icon: 'add_column_left' },
  'sheet.command.insert-col-before': { label: 'Inserted column left', icon: 'add_column_left' },
  'sheet.command.insert-col-after': { label: 'Inserted column right', icon: 'add_column_right' },
  'sheet.command.remove-row': { label: 'Deleted row(s)', icon: 'delete' },
  'sheet.command.remove-col': { label: 'Deleted column(s)', icon: 'delete' },
  'sheet.command.delete-range-move-up': { label: 'Deleted cells', icon: 'delete' },
  'sheet.command.delete-range-move-left': { label: 'Deleted cells', icon: 'delete' },
  'sheet.command.insert-range-move-down': { label: 'Inserted cells', icon: 'add' },
  'sheet.command.insert-range-move-right': { label: 'Inserted cells', icon: 'add' },
  'sheet.command.move-range': { label: 'Moved cells', icon: 'open_with' },
  'sheet.command.move-rows': { label: 'Moved row(s)', icon: 'swap_vert' },
  'sheet.command.move-cols': { label: 'Moved column(s)', icon: 'swap_horiz' },
  'sheet.command.add-worksheet-merge-all': { label: 'Merged cells', icon: 'cell_merge' },
  'sheet.command.remove-worksheet-merge': { label: 'Unmerged cells', icon: 'grid_on' },
  'sheet.command.set-worksheet-row-height': { label: 'Resized row(s)', icon: 'height' },
  'sheet.command.set-worksheet-col-width': { label: 'Resized column(s)', icon: 'width_normal' },
  'sheet.command.set-row-hidden': { label: 'Hid row(s)', icon: 'visibility_off' },
  'sheet.command.set-specific-rows-visible': { label: 'Showed row(s)', icon: 'visibility' },
  'sheet.command.set-col-hidden': { label: 'Hid column(s)', icon: 'visibility_off' },
  'sheet.command.set-specific-cols-visible': { label: 'Showed column(s)', icon: 'visibility' },
  'sheet.command.set-frozen': { label: 'Changed freeze panes', icon: 'ac_unit' },
  'sheet.command.insert-sheet': { label: 'Added a sheet', icon: 'add_box' },
  'sheet.command.remove-sheet': { label: 'Removed a sheet', icon: 'delete' },
  'sheet.command.set-worksheet-name': {
    label: 'Renamed a sheet',
    icon: 'drive_file_rename_outline',
  },
  'sheet.command.set-worksheet-order': { label: 'Reordered sheets', icon: 'reorder' },
  'sheet.command.set-tab-color': { label: 'Changed tab color', icon: 'palette' },
  'sheet.command.paste': { label: 'Pasted', icon: 'content_paste' },
  'sheet.command.copy': { label: 'Copied', icon: 'content_copy' },
  'sheet.command.cut': { label: 'Cut', icon: 'content_cut' },
  'sheet.command.add-table': { label: 'Inserted a table', icon: 'table' },
  'sheet.command.delete-table': { label: 'Deleted a table', icon: 'delete' },
  'sheet.command.numfmt.set.numfmt': { label: 'Changed number format', icon: 'tag' },
  'univer.command.undo': { label: 'Undo', icon: 'undo' },
  'univer.command.redo': { label: 'Redo', icon: 'redo' },
};

/**
 * Which command ids count as "user-facing activity". We only record
 * `CommandType.COMMAND` (intent) executions, and we skip pure selection /
 * navigation / scroll operations even though some of those are dispatched as
 * COMMANDs — they're noise, not edits. The allow-list above IS the primary
 * signal; unmapped COMMANDs still record with a fallback label unless they
 * match an obvious noise prefix.
 */
const NOISE_PREFIXES = [
  'sheet.operation.',
  'sheet.command.set-selection',
  'sheet.command.scroll',
  'sheet.command.set-zoom',
  'sheet.command.set-activate-cell-edit',
  'sheet.command.set-cell-edit-visible',
  'sheet.command.set-editor',
  'formula.',
  'univer.command.set-current-locale',
];

function isNoise(id: string): boolean {
  return NOISE_PREFIXES.some((p) => id.startsWith(p));
}

function labelFor(id: string): { label: string; icon: string } {
  const known = COMMAND_LABELS[id];
  if (known) return known;
  // De-namespace the id into something human-ish.
  const tail = id
    .replace(/^sheet\.command\./, '')
    .replace(/^univer\.command\./, '')
    .replace(/[.-]/g, ' ')
    .trim();
  return { label: tail ? tail.charAt(0).toUpperCase() + tail.slice(1) : 'Change', icon: 'bolt' };
}

function relativeTime(t: number, now: number): string {
  const delta = now - t;
  if (delta < 5_000) return 'just now';
  if (delta < 60_000) return `${Math.floor(delta / 1000)}s ago`;
  if (delta < 3_600_000) return `${Math.floor(delta / 60_000)}m ago`;
  return `${Math.floor(delta / 3_600_000)}h ago`;
}

const headerStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '10px 12px',
  borderBottom: '1px solid var(--cs-chrome-border, #edeff3)',
};

export function HistoryPanel({ api, onClose }: PanelComponentProps) {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const univer = (api as any).univer;
  const [entries, setEntries] = useState<ActivityEntry[]>([]);
  // Ordering counter — monotonic, initialised at 0 (no wall-clock at load).
  const seqRef = useRef(0);
  // Re-render tick so relative times refresh even without new activity.
  const [nowTick, setNowTick] = useState(0);

  useEffect(() => {
    if (!univer) return;
    const injector = (univer as { _injector?: { get(t: unknown): unknown } })._injector;
    const cmd = injector?.get(ICommandService) as
      | {
          onCommandExecuted: (cb: (info: ICommandInfo, options?: unknown) => void) => {
            dispose: () => void;
          };
        }
      | undefined;
    if (!cmd) return;

    const sub = cmd.onCommandExecuted((info) => {
      // Only intent-level COMMANDs; mutations/operations would spam the feed.
      // `type` is optional on ICommandInfo but the dispatcher populates it —
      // treat an explicit non-COMMAND type as a reject, and let undefined
      // fall through to the id-based allow-list below.
      if (info.type !== undefined && info.type !== CommandType.COMMAND) return;
      if (isNoise(info.id)) return;
      // Ignore anything not on the allow-list unless it looks like a real
      // sheet command (keeps the feed to genuine edits, not internal churn).
      if (!COMMAND_LABELS[info.id] && !info.id.startsWith('sheet.command.')) return;

      const { label, icon } = labelFor(info.id);
      const seq = ++seqRef.current;
      const entry: ActivityEntry = { seq, id: info.id, label, icon, t: Date.now() };
      setEntries((prev) => {
        const next = prev.length >= RING_CAP ? prev.slice(prev.length - RING_CAP + 1) : prev;
        return [...next, entry];
      });
    });
    return () => sub.dispose();
  }, [univer]);

  // Refresh relative timestamps once a minute while the panel is open.
  useEffect(() => {
    const h = setInterval(() => setNowTick((n) => n + 1), 60_000);
    return () => clearInterval(h);
  }, []);

  // Newest-first. Ordering is purely by seq — the render-time `now` only
  // affects the human-readable age string, never the order.
  const sorted = useMemo(() => [...entries].sort((a, b) => b.seq - a.seq), [entries]);
  const now = useMemo(() => Date.now(), [nowTick, entries.length]);

  const empty = sorted.length === 0;

  return (
    <div
      data-testid="cs-history-panel"
      style={{ height: '100%', display: 'flex', flexDirection: 'column' }}
    >
      <PanelHeader icon="history" title="Activity" count={sorted.length} onClose={onClose} />

      <div style={{ flex: 1, overflow: 'auto', padding: 12 }}>
        {empty ? (
          <div
            data-testid="cs-history-panel-empty"
            style={{ textAlign: 'center', opacity: 0.75, padding: '24px 8px' }}
          >
            <Icon name="history" size={40} style={{ opacity: 0.4 }} />
            <div style={{ fontWeight: 600, marginTop: 8 }}>No activity yet</div>
            <div style={{ fontSize: 13, marginTop: 4 }}>
              Edits you make this session show up here.
            </div>
          </div>
        ) : (
          <ol
            role="list"
            style={{
              listStyle: 'none',
              margin: 0,
              padding: 0,
              display: 'flex',
              flexDirection: 'column',
              gap: 2,
            }}
          >
            {sorted.map((e) => (
              <li
                key={e.seq}
                data-testid="cs-history-row"
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: 10,
                  padding: '8px 8px',
                  borderRadius: 8,
                }}
              >
                <span
                  aria-hidden
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    width: 28,
                    height: 28,
                    flex: '0 0 auto',
                    borderRadius: 6,
                    background: 'var(--cs-chrome-subtle-bg, #f4f5f8)',
                    opacity: 0.9,
                  }}
                >
                  <Icon name={e.icon} size={16} />
                </span>
                <span style={{ flex: 1, minWidth: 0, fontSize: 13, fontWeight: 500 }}>
                  {e.label}
                </span>
                <time
                  dateTime={new Date(e.t).toISOString()}
                  style={{ opacity: 0.6, fontSize: 12, whiteSpace: 'nowrap' }}
                >
                  {relativeTime(e.t, now)}
                </time>
              </li>
            ))}
          </ol>
        )}
      </div>
    </div>
  );
}
