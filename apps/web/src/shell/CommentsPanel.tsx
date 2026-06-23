import { useEffect, useState } from 'react';
import { useUniverAPI } from '../use-univer';
import { useUI } from '../use-ui';
import { Icon } from './Icon';
import { showCommentModal } from './tab-actions';

/**
 * Comments task pane — our own React panel so it shares the exact
 * `.side-panel` shell (header, float, slide-in motion, empty state) with
 * Tables / Charts / Outline / History, instead of Univer's bespoke
 * "Comment Management" sidebar which docked inside the canvas with its own
 * structure + animation.
 *
 * It indexes the thread comments on the active sheet (via the facade) and
 * navigates to a cell on click — Univer's in-cell comment popup then owns
 * the thread (reply / resolve / edit). "Add comment" opens the same
 * comment-on-cell modal as Review → New comment.
 */
type CommentRow = { id: string; text: string; ref: string; replies: number };

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function readComments(api: any): CommentRow[] {
  const ws = api?.getActiveWorkbook?.()?.getActiveSheet?.();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const all: any[] = ws?.getComments?.() ?? [];
  const rows: CommentRow[] = [];
  for (const c of all) {
    try {
      if (c.getIsRoot && !c.getIsRoot()) continue; // replies render under their root
      const data = c.getCommentData?.() ?? {};
      const stream: string = data.text?.dataStream ?? '';
      const text = stream.replace(/[\r\n\t]+/g, ' ').trim() || '(empty comment)';
      let ref = '';
      try {
        ref = c.getRange?.()?.getA1Notation?.() ?? '';
      } catch {
        /* range no longer resolvable */
      }
      const replies = c.getReplies?.()?.length ?? 0;
      rows.push({ id: c.id ?? data.id ?? ref, text, ref, replies });
    } catch {
      /* skip malformed thread */
    }
  }
  return rows;
}

export function CommentsPanel() {
  const api = useUniverAPI();
  const ui = useUI();
  const [rows, setRows] = useState<CommentRow[]>([]);

  useEffect(() => {
    if (!api) return;
    const read = () => setRows(readComments(api));
    read();
    const disp = api.addEvent(api.Event.CommandExecuted, (e) => {
      const id = (e as { id?: string }).id;
      if (!id) return;
      if (
        id.includes('comment') ||
        id === 'sheet.operation.set-worksheet-activate' ||
        id === 'doc.command-replace-snapshot'
      ) {
        read();
      }
    });
    return () => disp.dispose();
  }, [api]);

  const empty = rows.length === 0;

  const openThread = (r: CommentRow) => {
    if (!api || !r.ref) return;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (api.getActiveWorkbook()?.getActiveSheet() as any)?.getRange?.(r.ref)?.activate?.();
    } catch {
      /* range gone */
    }
  };

  // Resolve a comment. The command + collab sync already exist — this surfaces
  // them in the panel. Resolving removes the comment from the cell-location
  // index, so it leaves the list (the active view shows open threads only;
  // a resolved-comments view + reopen is a follow-up — see #111).
  const resolveComment = (r: CommentRow) => {
    if (!api) return;
    try {
      const wb = api.getActiveWorkbook();
      const ws = wb?.getActiveSheet();
      if (!wb || !ws) return;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const subUnitId = (ws as any).getSheetId?.() ?? (ws as any).getId?.();
      api.executeCommand('thread-comment.command.resolve-comment', {
        unitId: wb.getId(),
        subUnitId,
        commentId: r.id,
        resolved: true,
      });
    } catch {
      /* command unavailable — plugin not loaded */
    }
  };

  return (
    <aside className="side-panel comments-panel" data-testid="comments-panel">
      <header className="side-panel__header">
        <Icon name="forum" size="sm" />
        <h2 className="side-panel__title">Comments</h2>
        {!empty && <span className="side-panel__count">{rows.length}</span>}
        <button
          type="button"
          className="side-panel__close"
          aria-label="Close comments panel"
          onClick={ui.toggleCommentsPanel}
        >
          <Icon name="close" size="sm" />
        </button>
      </header>
      <div className="side-panel__body">
        {empty ? (
          <div className="side-panel__empty" data-testid="comments-panel-empty">
            <Icon name="forum" size="lg" className="side-panel__empty-icon" />
            <div className="side-panel__empty-title">No comments yet</div>
            <div className="side-panel__empty-body">
              Select a cell and add a comment to start a discussion — or use{' '}
              <strong>Review → New comment</strong>.
            </div>
            <button
              type="button"
              className="btn-primary side-panel__empty-cta"
              data-testid="comments-panel-empty-cta"
              disabled={!api}
              onClick={() => api && showCommentModal(api)}
            >
              Add comment
            </button>
          </div>
        ) : (
          <ul className="side-panel__rows comments-panel__list">
            {rows.map((r) => (
              <li
                className="side-panel__row comments-panel__row"
                key={r.id}
                data-testid={`comments-panel-row-${r.id}`}
              >
                <button
                  type="button"
                  className="comments-panel__open"
                  onClick={() => openThread(r)}
                  title={`Go to ${r.ref || 'comment'}`}
                >
                  <span className="comments-panel__row-top">
                    <span className="comments-panel__ref">{r.ref || 'Comment'}</span>
                    {r.replies > 0 && (
                      <span className="comments-panel__replies" title={`${r.replies} replies`}>
                        <Icon name="reply" size="sm" />
                        {r.replies}
                      </span>
                    )}
                  </span>
                  <span className="comments-panel__text">{r.text}</span>
                </button>
                <button
                  type="button"
                  className="comments-panel__resolve"
                  data-testid={`comments-panel-resolve-${r.id}`}
                  title="Resolve comment"
                  onClick={() => resolveComment(r)}
                >
                  <Icon name="check_circle" size="sm" />
                </button>
              </li>
            ))}
          </ul>
        )}
      </div>
    </aside>
  );
}
