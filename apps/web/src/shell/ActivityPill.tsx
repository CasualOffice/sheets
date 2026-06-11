import { useEffect, useRef, useState } from 'react';
import { Icon } from './Icon';
import { useActivity } from './activity-context';

/**
 * Title-bar bell icon + popover for the activity log.
 * UX_AUDIT.md §4.1 / Phase 4 #14.
 *
 * - Idle (no entries): renders nothing — the bar stays clean.
 * - With entries: bell icon; a red badge shows the unread count.
 * - Click opens a small popover with the latest N errors. Each row
 *   has a Dismiss button. Footer has "Clear all".
 */
export function ActivityPill() {
  const { entries, unread, markAllRead, dismiss, clearAll } = useActivity();
  const [open, setOpen] = useState(false);
  const wrapRef = useRef<HTMLDivElement | null>(null);

  // Close on outside click + Escape.
  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target as Node)) setOpen(false);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(false);
    };
    document.addEventListener('mousedown', onDown);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('mousedown', onDown);
      document.removeEventListener('keydown', onKey);
    };
  }, [open]);

  if (entries.length === 0) return null;

  const toggle = () => {
    setOpen((prev) => {
      const next = !prev;
      if (next) markAllRead();
      return next;
    });
  };

  return (
    <div
      ref={wrapRef}
      data-testid="activity-pill"
      style={{ position: 'relative', display: 'inline-flex' }}
    >
      <button
        type="button"
        className="titlebar__icon-btn"
        aria-haspopup="menu"
        aria-expanded={open}
        aria-label={
          unread > 0
            ? `Activity log — ${unread} new error${unread === 1 ? '' : 's'}`
            : 'Activity log'
        }
        title="Activity log"
        onClick={toggle}
        data-testid="activity-pill-trigger"
        data-unread={unread > 0 ? '1' : '0'}
        style={{ position: 'relative' }}
      >
        <Icon name="bug_report" size="sm" />
        {unread > 0 && (
          <span
            aria-hidden="true"
            data-testid="activity-pill-badge"
            style={{
              position: 'absolute',
              top: 2,
              right: 2,
              minWidth: 14,
              height: 14,
              padding: '0 4px',
              borderRadius: 7,
              background: '#dc2626',
              color: '#fff',
              fontSize: 9,
              fontWeight: 700,
              lineHeight: '14px',
              textAlign: 'center',
              boxSizing: 'border-box',
            }}
          >
            {unread > 9 ? '9+' : unread}
          </span>
        )}
      </button>
      {open && (
        <div
          role="menu"
          data-testid="activity-pill-popover"
          style={{
            position: 'absolute',
            top: 'calc(100% + 6px)',
            right: 0,
            width: 320,
            maxHeight: 360,
            display: 'flex',
            flexDirection: 'column',
            background: '#fff',
            border: '1px solid #e2e8f0',
            borderRadius: 8,
            boxShadow:
              '0 1px 2px rgba(0,0,0,0.04), 0 8px 24px rgba(15,23,42,0.10)',
            zIndex: 1000,
            overflow: 'hidden',
          }}
        >
          <header
            style={{
              padding: '10px 12px',
              borderBottom: '1px solid #e2e8f0',
              fontSize: 12,
              fontWeight: 600,
              color: '#0f172a',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
            }}
          >
            <span>Activity</span>
            <span style={{ color: '#64748b', fontWeight: 400 }}>
              {entries.length} entr{entries.length === 1 ? 'y' : 'ies'}
            </span>
          </header>
          <ul
            style={{
              flex: 1,
              overflow: 'auto',
              listStyle: 'none',
              padding: 0,
              margin: 0,
            }}
          >
            {entries.map((entry) => (
              <li
                key={entry.id}
                style={{
                  padding: '10px 12px',
                  borderBottom: '1px solid #f1f5f9',
                  display: 'flex',
                  gap: 8,
                  alignItems: 'flex-start',
                }}
                data-testid={`activity-entry-${entry.id}`}
              >
                <Icon name="bug_report" size="sm" />
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div
                    style={{
                      fontSize: 12,
                      color: '#0f172a',
                      wordBreak: 'break-word',
                    }}
                  >
                    {entry.message}
                  </div>
                  <div
                    style={{ fontSize: 11, color: '#94a3b8', marginTop: 2 }}
                  >
                    {formatTime(entry.timestamp)}
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => dismiss(entry.id)}
                  aria-label="Dismiss"
                  data-testid={`activity-entry-${entry.id}-dismiss`}
                  style={{
                    border: 'none',
                    background: 'transparent',
                    cursor: 'pointer',
                    color: '#64748b',
                    padding: 0,
                    display: 'inline-flex',
                  }}
                >
                  <Icon name="close" size="sm" />
                </button>
              </li>
            ))}
          </ul>
          <footer
            style={{
              padding: '8px 12px',
              borderTop: '1px solid #e2e8f0',
              display: 'flex',
              justifyContent: 'flex-end',
            }}
          >
            <button
              type="button"
              onClick={() => {
                clearAll();
                setOpen(false);
              }}
              data-testid="activity-clear-all"
              style={{
                fontSize: 12,
                color: '#0f172a',
                background: 'transparent',
                border: 'none',
                cursor: 'pointer',
                padding: '4px 8px',
              }}
            >
              Clear all
            </button>
          </footer>
        </div>
      )}
    </div>
  );
}

function formatTime(ts: number): string {
  const d = new Date(ts);
  const now = Date.now();
  const diff = now - ts;
  if (diff < 60_000) return 'just now';
  if (diff < 3_600_000) return `${Math.round(diff / 60_000)} min ago`;
  return d.toLocaleTimeString();
}
