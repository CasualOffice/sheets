/**
 * MySpreadsheetsList — the personal-mode landing surface at /home.
 * UX_AUDIT.md §2.2, §2.5, §2.10.
 *
 * Single source of truth for "the user's workbooks". Subscribes to the
 * active FileSource's recent list, dedupes by filename (latest per
 * name), sorts modifiedAt DESC, renders one row per file with a kebab
 * menu (Rename / Delete / Download — Phase 1 ships just Open + Delete,
 * the rest plug in per UX_AUDIT.md Phase 2).
 *
 * The dedup is belt-and-braces — once the in-place save bug from
 * UX_AUDIT.md §2.3 lands, every save overwrites the same id so the
 * list is naturally dedupe'd. Until then this keeps the visible bug
 * gone.
 */

import { useEffect, useMemo, useRef, useState, type CSSProperties } from 'react';

import { useFileSource } from '../file-source';
import { navigate } from '../router';
import type { RecentEntry } from '../file-source/types';

interface ListState {
  status: 'idle' | 'loading' | 'ready' | 'error';
  entries: RecentEntry[];
  errorMessage?: string;
}

const INITIAL: ListState = { status: 'idle', entries: [] };

export function MySpreadsheetsList() {
  const fs = useFileSource();
  const [state, setState] = useState<ListState>(INITIAL);
  const aliveRef = useRef(true);

  // Initial load + subscription to file-source changes (the personal
  // source bumps the subscription after every save / delete).
  useEffect(() => {
    aliveRef.current = true;
    const reload = async () => {
      if (!aliveRef.current) return;
      setState((s) => ({ ...s, status: s.entries.length ? 'ready' : 'loading' }));
      try {
        const rows = await fs.listRecent();
        if (!aliveRef.current) return;
        setState({ status: 'ready', entries: rows });
      } catch (err) {
        if (!aliveRef.current) return;
        setState({
          status: 'error',
          entries: [],
          errorMessage: err instanceof Error ? err.message : String(err),
        });
      }
    };
    void reload();
    const unsub = fs.subscribeRecent(reload);
    return () => {
      aliveRef.current = false;
      unsub();
    };
  }, [fs]);

  // Dedup by name (latest per name) + sort modifiedAt DESC. Once the
  // in-place save fix lands, dedup will be a no-op; until then it
  // collapses the duplicate-row symptom users see today.
  const visibleRows = useMemo(() => {
    const latestByName = new Map<string, RecentEntry>();
    for (const row of state.entries) {
      const prev = latestByName.get(row.name);
      if (!prev || row.modifiedAt > prev.modifiedAt) {
        latestByName.set(row.name, row);
      }
    }
    return Array.from(latestByName.values()).sort((a, b) => b.modifiedAt - a.modifiedAt);
  }, [state.entries]);

  return (
    <div style={pageStyle}>
      <header style={headerStyle}>
        <div>
          <h1 style={titleStyle}>My Spreadsheets</h1>
          <p style={subtitleStyle}>
            {state.status === 'ready' && visibleRows.length > 0
              ? `${visibleRows.length} ${visibleRows.length === 1 ? 'file' : 'files'}`
              : 'Saved workbooks live here.'}
          </p>
        </div>
        <div style={actionsStyle}>
          <button
            type="button"
            onClick={() => navigate('/templates')}
            style={secondaryBtnStyle}
            data-testid="home-new-from-template"
          >
            From template
          </button>
          <button
            type="button"
            onClick={() => navigate('/sheet/new')}
            style={primaryBtnStyle}
            data-testid="home-new-blank"
          >
            + New blank
          </button>
        </div>
      </header>

      {state.status === 'loading' && <SkeletonGrid />}
      {state.status === 'error' && (
        <ErrorPanel
          message={state.errorMessage ?? 'Failed to load files'}
          onRetry={() => {
            // Toggle status to retrigger the useEffect via state churn.
            setState((s) => ({ ...s, status: 'idle' }));
          }}
        />
      )}
      {state.status === 'ready' && visibleRows.length === 0 && <EmptyState />}
      {state.status === 'ready' && visibleRows.length > 0 && (
        <FilesGrid
          rows={visibleRows}
          onOpen={(id) => navigate('/sheet/' + encodeURIComponent(id))}
          onDelete={async (id) => {
            try {
              await fs.forgetRecent(id);
              // subscribeRecent fires reload — no explicit setState needed.
            } catch (err) {
              const msg = err instanceof Error ? err.message : String(err);
              window.alert(`Couldn't delete: ${msg}`);
            }
          }}
        />
      )}
    </div>
  );
}

// ─── Subcomponents ──────────────────────────────────────────────────────

function FilesGrid({
  rows,
  onOpen,
  onDelete,
}: {
  rows: RecentEntry[];
  onOpen: (id: string) => void;
  onDelete: (id: string) => void;
}) {
  return (
    <ul style={gridStyle} data-testid="home-files-grid">
      {rows.map((row) => (
        <FileRow key={row.id} row={row} onOpen={onOpen} onDelete={onDelete} />
      ))}
    </ul>
  );
}

function FileRow({
  row,
  onOpen,
  onDelete,
}: {
  row: RecentEntry;
  onOpen: (id: string) => void;
  onDelete: (id: string) => void;
}) {
  const [hover, setHover] = useState(false);
  return (
    <li
      style={rowStyle(hover)}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      onClick={() => onOpen(row.id)}
      onKeyDown={(e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          onOpen(row.id);
        }
      }}
      role="button"
      tabIndex={0}
      data-testid={`home-file-row-${row.id}`}
    >
      <div style={thumbStyle} aria-hidden="true">
        <span style={thumbGlyphStyle}>xlsx</span>
      </div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={rowNameStyle}>{row.name}</div>
        <div style={rowMetaStyle}>
          {formatRelative(row.modifiedAt)} · {formatBytes(row.size)}
        </div>
      </div>
      <button
        type="button"
        style={deleteBtnStyle(hover)}
        onClick={(e) => {
          e.stopPropagation();
          if (window.confirm(`Delete "${row.name}"?`)) {
            onDelete(row.id);
          }
        }}
        aria-label={`Delete ${row.name}`}
        data-testid={`home-file-delete-${row.id}`}
      >
        Delete
      </button>
    </li>
  );
}

function SkeletonGrid() {
  return (
    <ul style={gridStyle} aria-hidden="true">
      {Array.from({ length: 4 }, (_, i) => (
        <li key={i} style={skeletonRowStyle}>
          <div style={skeletonThumbStyle} />
          <div style={{ flex: 1 }}>
            <div style={skeletonLineStyle(60)} />
            <div style={{ ...skeletonLineStyle(40), marginTop: 8 }} />
          </div>
        </li>
      ))}
    </ul>
  );
}

function EmptyState() {
  return (
    <div style={emptyWrapStyle} data-testid="home-empty">
      <div style={emptyIconStyle} aria-hidden="true">
        xlsx
      </div>
      <h2 style={emptyTitleStyle}>No spreadsheets yet</h2>
      <p style={emptySubStyle}>Create your first workbook or pick a template to get started.</p>
      <div style={{ display: 'flex', gap: 10, marginTop: 18 }}>
        <button
          type="button"
          onClick={() => navigate('/templates')}
          style={secondaryBtnStyle}
        >
          Browse templates
        </button>
        <button type="button" onClick={() => navigate('/sheet/new')} style={primaryBtnStyle}>
          + New blank
        </button>
      </div>
    </div>
  );
}

function ErrorPanel({ message, onRetry }: { message: string; onRetry: () => void }) {
  return (
    <div style={errorWrapStyle}>
      <h2 style={emptyTitleStyle}>Couldn&apos;t load files</h2>
      <p style={emptySubStyle}>{message}</p>
      <button type="button" onClick={onRetry} style={primaryBtnStyle}>
        Retry
      </button>
    </div>
  );
}

// ─── Format helpers ─────────────────────────────────────────────────────

function formatBytes(n: number): string {
  if (!Number.isFinite(n) || n <= 0) return '—';
  const units = ['B', 'KB', 'MB', 'GB'];
  let v = n;
  let i = 0;
  while (v >= 1024 && i < units.length - 1) {
    v /= 1024;
    i += 1;
  }
  return `${i === 0 ? v : v.toFixed(v < 10 ? 1 : 0)} ${units[i]}`;
}

function formatRelative(ts: number): string {
  if (!Number.isFinite(ts) || ts <= 0) return '—';
  const now = Date.now();
  const diff = Math.max(0, now - ts);
  const sec = Math.round(diff / 1000);
  if (sec < 60) return 'just now';
  const min = Math.round(sec / 60);
  if (min < 60) return `${min} min ago`;
  const hr = Math.round(min / 60);
  if (hr < 24) return `${hr} hr ago`;
  const day = Math.round(hr / 24);
  if (day < 7) return `${day} day${day === 1 ? '' : 's'} ago`;
  return new Date(ts).toLocaleDateString(undefined, {
    month: 'short',
    day: 'numeric',
    year:
      new Date(ts).getFullYear() === new Date(now).getFullYear() ? undefined : 'numeric',
  });
}

// ─── Styles (inline; mirrors HomeScreen.tsx's idiom) ────────────────────

const pageStyle: CSSProperties = {
  maxWidth: 880,
  margin: '0 auto',
  padding: '40px 24px 80px',
  display: 'flex',
  flexDirection: 'column',
  gap: 24,
  fontFamily:
    'system-ui, -apple-system, "Segoe UI", "Helvetica Neue", Arial, sans-serif',
};

const headerStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'flex-end',
  justifyContent: 'space-between',
  gap: 16,
  flexWrap: 'wrap',
};

const titleStyle: CSSProperties = {
  margin: 0,
  fontSize: 28,
  fontWeight: 600,
  letterSpacing: '-0.01em',
  color: '#0f172a',
};

const subtitleStyle: CSSProperties = {
  margin: '6px 0 0',
  fontSize: 14,
  color: '#64748b',
};

const actionsStyle: CSSProperties = { display: 'flex', gap: 8 };

const primaryBtnStyle: CSSProperties = {
  padding: '9px 16px',
  borderRadius: 8,
  border: '1px solid #15803d',
  background: '#15803d',
  color: '#fff',
  fontSize: 14,
  fontWeight: 500,
  cursor: 'pointer',
};

const secondaryBtnStyle: CSSProperties = {
  padding: '9px 16px',
  borderRadius: 8,
  border: '1px solid #cbd5e1',
  background: '#fff',
  color: '#0f172a',
  fontSize: 14,
  fontWeight: 500,
  cursor: 'pointer',
};

const gridStyle: CSSProperties = {
  listStyle: 'none',
  padding: 0,
  margin: 0,
  display: 'flex',
  flexDirection: 'column',
  gap: 6,
};

function rowStyle(hover: boolean): CSSProperties {
  return {
    display: 'flex',
    alignItems: 'center',
    gap: 14,
    padding: '14px 16px',
    background: hover ? '#f1f5f9' : '#fff',
    border: '1px solid #e2e8f0',
    borderRadius: 10,
    cursor: 'pointer',
    transition: 'background 120ms ease',
  };
}

const thumbStyle: CSSProperties = {
  width: 44,
  height: 44,
  borderRadius: 8,
  background: '#dcfce7',
  color: '#166534',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  fontSize: 11,
  fontWeight: 600,
  letterSpacing: 0.5,
  flexShrink: 0,
};

const thumbGlyphStyle: CSSProperties = { textTransform: 'uppercase' };

const rowNameStyle: CSSProperties = {
  fontSize: 14,
  fontWeight: 500,
  color: '#0f172a',
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap',
};

const rowMetaStyle: CSSProperties = {
  fontSize: 12,
  color: '#64748b',
  marginTop: 2,
};

function deleteBtnStyle(hover: boolean): CSSProperties {
  return {
    padding: '6px 12px',
    borderRadius: 6,
    border: '1px solid #fecaca',
    background: hover ? '#fef2f2' : '#fff',
    color: '#b91c1c',
    fontSize: 12,
    fontWeight: 500,
    cursor: 'pointer',
    visibility: hover ? 'visible' : 'hidden',
  };
}

const skeletonRowStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 14,
  padding: '14px 16px',
  background: '#fff',
  border: '1px solid #e2e8f0',
  borderRadius: 10,
};

const skeletonThumbStyle: CSSProperties = {
  width: 44,
  height: 44,
  borderRadius: 8,
  background: '#e2e8f0',
  flexShrink: 0,
};

function skeletonLineStyle(widthPct: number): CSSProperties {
  return {
    width: `${widthPct}%`,
    height: 12,
    background: '#e2e8f0',
    borderRadius: 4,
  };
}

const emptyWrapStyle: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  textAlign: 'center',
  padding: '64px 24px',
  background: '#fff',
  border: '1px dashed #cbd5e1',
  borderRadius: 16,
};

const emptyIconStyle: CSSProperties = {
  width: 56,
  height: 56,
  borderRadius: 12,
  background: '#dcfce7',
  color: '#166534',
  fontSize: 12,
  fontWeight: 600,
  letterSpacing: 0.5,
  textTransform: 'uppercase',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  marginBottom: 16,
};

const emptyTitleStyle: CSSProperties = {
  margin: '0 0 6px',
  fontSize: 18,
  fontWeight: 600,
  color: '#0f172a',
};

const emptySubStyle: CSSProperties = {
  margin: 0,
  fontSize: 14,
  color: '#64748b',
  maxWidth: 360,
};

const errorWrapStyle: CSSProperties = {
  ...emptyWrapStyle,
  borderColor: '#fecaca',
  background: '#fef2f2',
};
