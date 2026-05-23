import type { IWorkbookData } from '@univerjs/core';
import { useLiveRecentFiles } from '../recent-files/useLiveRecentFiles';
import { deleteRecentFile, type RecentFile } from '../recent-files/store';
import { useWorkbook } from '../use-workbook';
import { Icon } from './Icon';

/**
 * Empty-state landing rendered over the grid while the workbook is a
 * fresh blank `Untitled` and there are recent files to surface. Click
 * a card to reopen — it calls `replaceWorkbook` with the cached
 * snapshot, same path File→Open uses.
 *
 * Auto-hides as soon as the workbook becomes non-blank, so it doesn't
 * cover a workbook the user has started editing.
 */

export function RecentFilesLanding() {
  const wb = useWorkbook();
  const files = useLiveRecentFiles();

  // Show only on a blank Untitled and only when there's actually
  // something to show. The `meta.revision === 0` check means the user
  // hasn't loaded or seriously edited anything in this session.
  const isBlank = wb.meta.name === 'Untitled' && wb.meta.revision <= 1;
  if (!isBlank) return null;
  if (files.length === 0) return null;

  const onOpen = (rec: RecentFile) => {
    wb.replaceWorkbook(rec.data as IWorkbookData, rec.sourceFormat);
  };
  const onDelete = (rec: RecentFile) => {
    if (rec.id == null) return;
    void deleteRecentFile(rec.id);
  };

  return (
    <aside
      className="recent-files-landing"
      data-testid="recent-files-landing"
      aria-label="Recent files"
    >
      <header className="recent-files-landing__header">
        <Icon name="history" />
        <h2 className="recent-files-landing__title">Recent files</h2>
      </header>
      <ul className="recent-files-landing__list">
        {files.map((rec) => (
          <li
            key={rec.id}
            className="recent-files-landing__row"
            data-testid="recent-files-row"
          >
            <button
              type="button"
              className="recent-files-landing__open"
              data-testid="recent-files-open"
              onClick={() => onOpen(rec)}
            >
              <span className="recent-files-landing__icon">
                <Icon name="description" />
              </span>
              <span className="recent-files-landing__main">
                <span className="recent-files-landing__name">{rec.name}</span>
                <span className="recent-files-landing__meta">
                  {formatSize(rec.size)} · {formatTime(rec.openedAt)}
                </span>
              </span>
            </button>
            <button
              type="button"
              className="recent-files-landing__delete"
              data-testid="recent-files-delete"
              title="Remove from recent"
              onClick={() => onDelete(rec)}
            >
              <Icon name="delete" size="sm" />
            </button>
          </li>
        ))}
      </ul>
      <p className="recent-files-landing__hint">
        Drop an Excel / ODS file anywhere, or use{' '}
        <strong>File → Open</strong>.
      </p>
    </aside>
  );
}

function formatSize(bytes: number): string {
  if (!bytes) return '';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function formatTime(ts: number): string {
  const ms = Date.now() - ts;
  if (ms < 60_000) return 'just now';
  if (ms < 60 * 60_000) return `${Math.floor(ms / 60_000)} min ago`;
  if (ms < 24 * 60 * 60_000) return `${Math.floor(ms / 3_600_000)} hr ago`;
  const days = Math.floor(ms / 86_400_000);
  return `${days} day${days === 1 ? '' : 's'} ago`;
}
