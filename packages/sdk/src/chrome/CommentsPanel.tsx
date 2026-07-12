/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License").
 */

/**
 * Comments side panel: lists the thread comments on the active sheet (active +
 * resolved), navigates to a comment's cell on click, and resolves / reopens a
 * thread inline. "Add comment" opens Univer's comment-on-cell modal (the same
 * one as Review → New comment); the in-cell popup owns full reply threading.
 *
 * Ported from the standalone app's `shell/CommentsPanel`. The app-only collab
 * stores it depended on — `collab/comment-authors` (avatar authorship),
 * `collab/presence` (display name / initials) and `collab/comment-mentions`
 * (@-mention "mentions you" badge) — DO NOT exist in the SDK and are dropped:
 * the byline degrades to the in-band `personId` string carried on the comment
 * model (no avatar, no color), and the mentions-me highlight is omitted. The
 * `useUniverAPI()` / `useUI()` couplings are replaced by the `api` / `onClose`
 * panel-host props, and the Univer facade is reached through `api.univer`.
 */
import { useEffect, useMemo, useState, type CSSProperties } from 'react';
import { SheetsThreadCommentModel } from '@univerjs/sheets-thread-comment';

import { ensurePluginByName } from '../univer';
import type { PanelComponentProps } from './extensions';
import { Icon } from './Icon';

interface CommentRow {
  id: string;
  text: string;
  ref: string;
  author: string;
  replies: number;
}

function cleanText(stream: string | undefined): string {
  return (stream ?? '').replace(/[\r\n\t]+/g, ' ').trim() || '(empty comment)';
}

/**
 * Active comments come straight off the facade (`getComments()` returns the
 * cell-location index; resolved threads leave it, so they're read separately).
 * Only root comments are listed — replies render under their root in the
 * in-cell popup.
 */
/* eslint-disable @typescript-eslint/no-explicit-any */
function readActive(univer: any): CommentRow[] {
  const ws = univer?.getActiveWorkbook?.()?.getActiveSheet?.();
  const all: any[] = ws?.getComments?.() ?? [];
  const rows: CommentRow[] = [];
  for (const c of all) {
    try {
      if (c.getIsRoot && !c.getIsRoot()) continue;
      const data = c.getCommentData?.() ?? {};
      let ref = '';
      try {
        ref = c.getRange?.()?.getA1Notation?.() ?? '';
      } catch {
        /* range no longer resolvable */
      }
      rows.push({
        id: c.id ?? data.id ?? ref,
        text: cleanText(data.text?.dataStream),
        ref,
        author: data.personId ?? '',
        replies: c.getReplies?.()?.length ?? 0,
      });
    } catch {
      /* skip malformed thread */
    }
  }
  return rows;
}

/**
 * Resolved comments have left the cell-location index, so `getComments()` no
 * longer returns them. Read them straight from the model — the only path to a
 * "resolved" view + reopen.
 */
function readResolved(univer: any): CommentRow[] {
  try {
    const injector = (univer as { _injector?: { get(t: unknown): unknown } })?._injector;
    const model = injector?.get?.(SheetsThreadCommentModel) as any;
    const wb = univer?.getActiveWorkbook?.();
    const ws = wb?.getActiveSheet?.();
    if (!model?.getSubUnitAll || !wb || !ws) return [];
    const unitId = wb.getId();
    const subUnitId = ws.getSheetId?.() ?? ws.getId?.();
    const all: any[] = model.getSubUnitAll(unitId, subUnitId) ?? [];
    const rows: CommentRow[] = [];
    for (const c of all) {
      if (!c?.resolved) continue;
      rows.push({
        id: c.id,
        text: cleanText(c.text?.dataStream),
        ref: c.ref ?? '',
        author: c.personId ?? '',
        replies: (c.children ?? []).length,
      });
    }
    return rows;
  } catch {
    return [];
  }
}
/* eslint-enable @typescript-eslint/no-explicit-any */

const header: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '10px 12px',
  borderBottom: '1px solid var(--cs-chrome-border, #edeff3)',
};

const rowCard: CSSProperties = {
  border: '1px solid var(--cs-chrome-border, #edeff3)',
  borderRadius: 8,
  padding: 10,
  display: 'flex',
  alignItems: 'flex-start',
  gap: 6,
};

const iconBtn: CSSProperties = {
  border: 'none',
  background: 'transparent',
  cursor: 'pointer',
  color: 'inherit',
  padding: 0,
  display: 'inline-flex',
  alignItems: 'center',
};

function CommentCard({
  row,
  resolved,
  onOpen,
  onToggle,
}: {
  row: CommentRow;
  resolved: boolean;
  onOpen: (r: CommentRow) => void;
  onToggle: (id: string, value: boolean) => void;
}) {
  return (
    <li
      data-testid={`cs-comments-panel-${resolved ? 'resolved-row' : 'row'}-${row.id}`}
      style={{ ...rowCard, opacity: resolved ? 0.72 : 1 }}
    >
      <button
        type="button"
        onClick={() => onOpen(row)}
        title={`Go to ${row.ref || 'comment'}`}
        style={{
          flex: 1,
          minWidth: 0,
          textAlign: 'left',
          border: 'none',
          background: 'transparent',
          cursor: 'pointer',
          color: 'inherit',
          font: 'inherit',
          display: 'flex',
          flexDirection: 'column',
          gap: 4,
        }}
      >
        <span style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <span style={{ fontWeight: 600 }}>{row.ref || 'Comment'}</span>
          {row.author && <span style={{ opacity: 0.6, fontSize: 12 }}>{row.author}</span>}
          {row.replies > 0 && (
            <span
              title={`${row.replies} replies`}
              style={{
                opacity: 0.6,
                fontSize: 12,
                display: 'inline-flex',
                alignItems: 'center',
                gap: 2,
              }}
            >
              <Icon name="reply" size={13} />
              {row.replies}
            </span>
          )}
        </span>
        <span
          style={{
            fontSize: 13,
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            display: '-webkit-box',
            WebkitLineClamp: 2,
            WebkitBoxOrient: 'vertical',
          }}
        >
          {row.text}
        </span>
      </button>
      <button
        type="button"
        data-testid={`cs-comments-panel-${resolved ? 'reopen' : 'resolve'}-${row.id}`}
        aria-label={resolved ? 'Reopen comment' : 'Resolve comment'}
        title={resolved ? 'Reopen comment' : 'Resolve comment'}
        onClick={() => onToggle(row.id, !resolved)}
        style={iconBtn}
      >
        <Icon name={resolved ? 'undo' : 'check_circle'} size={16} />
      </button>
    </li>
  );
}

const REFRESH_CMDS = (id: string) =>
  id.includes('comment') ||
  id === 'sheet.operation.set-worksheet-activate' ||
  id === 'doc.command-replace-snapshot';

export function CommentsPanel({ api, onClose }: PanelComponentProps) {
  const univer = (api as { univer?: any }).univer; // eslint-disable-line @typescript-eslint/no-explicit-any
  const [rows, setRows] = useState<CommentRow[]>([]);
  const [resolved, setResolved] = useState<CommentRow[]>([]);
  const [showResolved, setShowResolved] = useState(false);

  useEffect(() => {
    if (!univer) return;
    let cancelled = false;
    let disp: { dispose?: () => void } | undefined;
    const read = () => {
      if (cancelled) return;
      setRows(readActive(univer));
      setResolved(readResolved(univer));
    };
    // The thread-comment plugin is lazy — `getComments()` and
    // SheetsThreadCommentModel only resolve once its group is registered.
    void ensurePluginByName('threadComment').then(() => {
      if (cancelled) return;
      read();
      disp = univer.addEvent(univer.Event.CommandExecuted, (e: { id?: string }) => {
        if (e.id && REFRESH_CMDS(e.id)) read();
      });
    });
    return () => {
      cancelled = true;
      disp?.dispose?.();
    };
  }, [univer]);

  const empty = rows.length === 0 && resolved.length === 0;

  const openThread = (r: CommentRow) => {
    if (!univer || !r.ref) return;
    try {
      univer.getActiveWorkbook?.()?.getActiveSheet?.()?.getRange?.(r.ref)?.activate?.();
    } catch {
      /* range gone */
    }
  };

  const setResolvedState = async (commentId: string, value: boolean) => {
    if (!univer) return;
    try {
      const wb = univer.getActiveWorkbook?.();
      const ws = wb?.getActiveSheet?.();
      if (!wb || !ws) return;
      const subUnitId = ws.getSheetId?.() ?? ws.getId?.();
      await ensurePluginByName('threadComment');
      univer.executeCommand('thread-comment.command.resolve-comment', {
        unitId: wb.getId(),
        subUnitId,
        commentId,
        resolved: value,
      });
    } catch {
      /* command unavailable — plugin not loaded */
    }
  };

  const addComment = async () => {
    if (!univer) return;
    await ensurePluginByName('threadComment');
    univer.executeCommand('sheet.operation.show-comment-modal');
  };

  const handlers = useMemo(
    () => ({ open: openThread, toggle: setResolvedState }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [univer],
  );

  return (
    <div
      data-testid="cs-comments-panel"
      style={{ height: '100%', display: 'flex', flexDirection: 'column' }}
    >
      <header style={header}>
        <Icon name="forum" size={18} />
        <span style={{ fontWeight: 600, flex: 1 }}>Comments</span>
        {rows.length > 0 && <span style={{ opacity: 0.6, fontSize: 12 }}>{rows.length}</span>}
        <button
          type="button"
          aria-label="Add comment"
          title="Add comment"
          onClick={addComment}
          style={iconBtn}
        >
          <Icon name="add_comment" size={18} />
        </button>
        <button type="button" aria-label="Close comments panel" onClick={onClose} style={iconBtn}>
          <Icon name="close" size={18} />
        </button>
      </header>

      <div style={{ flex: 1, overflow: 'auto', padding: 12 }}>
        {empty ? (
          <div
            data-testid="cs-comments-panel-empty"
            style={{ textAlign: 'center', opacity: 0.75, padding: '24px 8px' }}
          >
            <Icon name="forum" size={40} style={{ opacity: 0.4 }} />
            <div style={{ fontWeight: 600, marginTop: 8 }}>No comments yet</div>
            <div style={{ fontSize: 13, marginTop: 4 }}>
              Select a cell and add a comment to start a discussion — or use{' '}
              <strong>Review → New comment</strong>.
            </div>
            <button
              type="button"
              data-testid="cs-comments-panel-empty-cta"
              disabled={!univer}
              onClick={addComment}
              style={{
                marginTop: 12,
                padding: '6px 14px',
                borderRadius: 6,
                border: 'none',
                cursor: univer ? 'pointer' : 'default',
                background: 'var(--cs-chrome-active-fg, #0e7490)',
                color: '#fff',
                font: 'inherit',
                fontWeight: 600,
              }}
            >
              Add comment
            </button>
          </div>
        ) : (
          <>
            {rows.length > 0 && (
              <ul
                style={{
                  listStyle: 'none',
                  margin: 0,
                  padding: 0,
                  display: 'flex',
                  flexDirection: 'column',
                  gap: 8,
                }}
              >
                {rows.map((r) => (
                  <CommentCard
                    key={r.id}
                    row={r}
                    resolved={false}
                    onOpen={handlers.open}
                    onToggle={handlers.toggle}
                  />
                ))}
              </ul>
            )}

            {resolved.length > 0 && (
              <div style={{ marginTop: rows.length > 0 ? 12 : 0 }}>
                <button
                  type="button"
                  data-testid="cs-comments-panel-resolved-toggle"
                  aria-expanded={showResolved}
                  onClick={() => setShowResolved((v) => !v)}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: 4,
                    width: '100%',
                    border: 'none',
                    background: 'transparent',
                    cursor: 'pointer',
                    color: 'inherit',
                    font: 'inherit',
                    fontWeight: 600,
                    padding: '4px 0',
                  }}
                >
                  <Icon name={showResolved ? 'expand_more' : 'chevron_right'} size={16} />
                  Resolved
                  <span style={{ opacity: 0.6, fontSize: 12 }}>{resolved.length}</span>
                </button>
                {showResolved && (
                  <ul
                    style={{
                      listStyle: 'none',
                      margin: '8px 0 0',
                      padding: 0,
                      display: 'flex',
                      flexDirection: 'column',
                      gap: 8,
                    }}
                  >
                    {resolved.map((r) => (
                      <CommentCard
                        key={r.id}
                        row={r}
                        resolved
                        onOpen={handlers.open}
                        onToggle={handlers.toggle}
                      />
                    ))}
                  </ul>
                )}
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
