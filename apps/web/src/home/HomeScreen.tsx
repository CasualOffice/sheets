import { useEffect, useMemo, useState } from 'react';
import type { IWorkbookData } from '@univerjs/core';
import { Icon } from '../shell/Icon';
import { useWorkbook } from '../use-workbook';
import { useLoading } from '../loading-context';
import { useFileSource, useRecentFiles, type RecentEntry } from '../file-source';
import { xlsxToWorkbookData } from '../xlsx';
import { emptyWorkbook } from '../snapshot';
import { loadSpreadsheetFile, pickXlsxFile } from '../shell/file-actions';
import { CATEGORIES, TEMPLATES, type Template, type TemplateCategory } from './registry';
import { TemplateCard } from './TemplateCard';
import { ReopenBanner } from './ReopenBanner';
import { IdbQuotaBanner } from './IdbQuotaBanner';
import { PinnedFolderSection } from './PinnedFolderSection';
import { usePinnedFolder } from '../file-system-access/usePinnedFolder';
import './home.css';

/**
 * The first thing a user sees when they land with an empty workbook —
 * a curated template gallery + their recent files. Replaces the older
 * `RecentFilesLanding` overlay: same gate condition (Untitled + early
 * revision), but the surface is now a full home page with hero + grid
 * instead of a small floating card.
 *
 * Templates are real .xlsx files under `public/templates/` — picking a
 * card runs the standard xlsx parse worker and swaps the workbook in
 * place. The overlay self-dismisses because the workbook is no longer
 * blank after a successful pick.
 */
export function HomeScreen({
  dismissed,
  onDismiss,
}: {
  dismissed: boolean;
  onDismiss: () => void;
}) {
  const wb = useWorkbook();
  const loading = useLoading();
  const fileSource = useFileSource();
  const recents = useRecentFiles();
  const pinned = usePinnedFolder();

  const isBlank = wb.meta.name === 'Untitled' && wb.meta.revision <= 1;
  // Suppress the home in collab rooms — the URL is the authoritative
  // signal here (collab context hydrates async, but Hocuspocus will
  // shortly replace the workbook with the room's snapshot and the home
  // would just flash). Matches CollabDriver's readRoomFromLocation():
  //   /r/:id           — canonical room URL
  //   /?room=:id        — query fallback
  // Same pattern the e2e prod-bundle specs use to enter rooms.
  const inCollabRoom =
    typeof window !== 'undefined' &&
    (/^\/r\/[\w-]{4,}\/?$/.test(window.location.pathname) ||
      (new URLSearchParams(window.location.search).get('room') ?? '').length >= 4);

  const [query, setQuery] = useState('');
  const [activeCategory, setActiveCategory] = useState<TemplateCategory | 'All'>('All');

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return TEMPLATES.filter((t) => {
      if (activeCategory !== 'All' && t.category !== activeCategory && t.id !== 'blank')
        return false;
      if (!q) return true;
      return (
        t.name.toLowerCase().includes(q) ||
        t.category.toLowerCase().includes(q) ||
        t.description.toLowerCase().includes(q)
      );
    });
  }, [query, activeCategory]);

  const featured = useMemo(() => TEMPLATES.filter((t) => t.featured), []);
  const byCategory = useMemo(() => {
    const map = new Map<TemplateCategory, Template[]>();
    for (const cat of CATEGORIES) map.set(cat, []);
    for (const t of TEMPLATES) {
      if (t.id === 'blank') continue; // featured-only
      map.get(t.category)?.push(t);
    }
    return map;
  }, []);

  const visible = !dismissed && isBlank && !inCollabRoom;

  // Esc closes the home, same as the X button.
  useEffect(() => {
    if (!visible) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        onDismiss();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [visible, onDismiss]);

  // Show on a blank Untitled workbook, until the user picks something
  // (template, Blank, or explicit close). Once dismissed it stays gone
  // for this session — re-entry will land via File → New later.
  if (!visible) return null;

  const pick = async (t: Template) => {
    if (t.id === 'blank') {
      // Blank pick keeps the workbook Untitled; explicit dismiss flips
      // the gate so the home goes away.
      wb.replaceWorkbook(emptyWorkbook(), null);
      onDismiss();
      return;
    }
    const url = `${import.meta.env.BASE_URL ?? '/'}templates/${t.id}.xlsx`;
    loading.set({ fileName: `${t.name}.xlsx`, phase: 'reading' });
    try {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`Template fetch failed: ${res.status}`);
      const sizeBytes = Number(res.headers.get('content-length') ?? '0') || undefined;
      loading.set({ phase: 'parsing', sizeBytes });
      const buf = await res.arrayBuffer();
      const data = (await xlsxToWorkbookData(buf)) as IWorkbookData;
      data.name = t.name;
      loading.set({ phase: 'mounting' });
      wb.replaceWorkbook(data, 'xlsx');
      onDismiss();
      // LoadingOverlay closes itself when the workbook swap completes.
      loading.set(null);
    } catch (err) {
      console.error('[home] template open failed', err);
      loading.set({
        fileName: `${t.name}.xlsx`,
        phase: 'reading',
        error: err instanceof Error ? err.message : 'Failed to open template.',
        onRetry: () => void pick(t),
      });
    }
  };

  const onOpenRecent = async (rec: RecentEntry) => {
    try {
      const opened = await fileSource.openRecent(rec.id);
      wb.replaceWorkbook(opened.data, opened.sourceFormat);
      onDismiss();
    } catch (err) {
      console.warn('[home] reopen failed', rec.id, err);
      loading.set({
        fileName: rec.name,
        phase: 'reading',
        error: err instanceof Error ? err.message : 'Could not reopen this file.',
      });
    }
  };

  const openFileFromDisk = async () => {
    const file = await pickXlsxFile();
    if (!file) return;
    loading.set({ fileName: file.name, sizeBytes: file.size, phase: 'reading' });
    try {
      await loadSpreadsheetFile(file, null, wb.replaceWorkbook, (phase) => loading.set({ phase }));
      onDismiss();
      loading.set(null);
    } catch (err) {
      console.error('[home] open file failed', err);
      loading.set({
        fileName: file.name,
        phase: 'reading',
        error: err instanceof Error ? err.message : 'Could not open this file.',
      });
    }
  };
  const onDeleteRecent = (rec: RecentEntry) => {
    void fileSource.forgetRecent(rec.id);
  };

  return (
    <div className="home" data-testid="home-screen" aria-label="Start a new spreadsheet">
      <button
        type="button"
        className="home__close"
        aria-label="Close home"
        title="Close (Esc)"
        onClick={onDismiss}
        data-testid="home-close"
      >
        <Icon name="close" />
      </button>
      <div className="home__scroll">
        <IdbQuotaBanner />
        <ReopenBanner recents={recents} onOpen={onOpenRecent} />
        <header className="home__hero">
          <div className="home__hero-glow" aria-hidden />
          <div className="home__hero-inner">
            <div className="home__brand">
              <Icon name="grid_on" />
              <span>Casual Sheets</span>
            </div>
            <h1 className="home__title">Start something today.</h1>
            <p className="home__lede">
              Pick a template designed for the way you actually work — or open a file from your
              computer.
            </p>
            <div className="home__hero-actions">
              <div className="home__search">
                <Icon name="search" />
                <input
                  type="search"
                  placeholder="Search templates"
                  value={query}
                  onChange={(e) => setQuery(e.target.value)}
                  data-testid="home-search"
                />
              </div>
              <button
                type="button"
                className="home__open-file"
                onClick={() => void openFileFromDisk()}
                data-testid="home-open-file"
              >
                <Icon name="folder_open" />
                <span>Open file</span>
              </button>
              <PinFolderControl
                state={pinned.state}
                onPin={() => void pinned.pin()}
                onReconnect={() => void pinned.reconnect()}
                onUnpin={() => void pinned.unpin()}
              />

              <div className="home__cats">
                <button
                  type="button"
                  className={`home__cat${activeCategory === 'All' ? ' home__cat--on' : ''}`}
                  onClick={() => setActiveCategory('All')}
                >
                  All
                </button>
                {CATEGORIES.map((c) => (
                  <button
                    key={c}
                    type="button"
                    className={`home__cat${activeCategory === c ? ' home__cat--on' : ''}`}
                    onClick={() => setActiveCategory(c)}
                  >
                    {c}
                  </button>
                ))}
              </div>
            </div>
          </div>
        </header>

        {query.trim() === '' && activeCategory === 'All' && (
          <section className="home__section home__section--featured">
            <div className="home__section-head">
              <h2>Featured</h2>
              <span className="home__section-hint">A few picks to get going.</span>
            </div>
            <div className="home__featured-row">
              {featured.map((t) => (
                <TemplateCard key={t.id} template={t} size="lg" onPick={pick} />
              ))}
            </div>
          </section>
        )}

        {query.trim() !== '' || activeCategory !== 'All' ? (
          <section className="home__section">
            <div className="home__section-head">
              <h2>{query.trim() ? `Results for "${query.trim()}"` : activeCategory}</h2>
              <span className="home__section-hint">
                {filtered.length} template{filtered.length === 1 ? '' : 's'}
              </span>
            </div>
            {filtered.length === 0 ? (
              <div className="home__empty">No templates match. Try a different keyword.</div>
            ) : (
              <div className="home__grid">
                {filtered.map((t) => (
                  <TemplateCard key={t.id} template={t} onPick={pick} />
                ))}
              </div>
            )}
          </section>
        ) : (
          <>
            {CATEGORIES.map((cat) => {
              const items = byCategory.get(cat) ?? [];
              if (items.length === 0) return null;
              return (
                <section key={cat} className="home__section">
                  <div className="home__section-head">
                    <h2>{cat}</h2>
                    <span className="home__section-hint">{items.length} templates</span>
                  </div>
                  <div className="home__grid">
                    {items.map((t) => (
                      <TemplateCard key={t.id} template={t} onPick={pick} />
                    ))}
                  </div>
                </section>
              );
            })}
          </>
        )}

        {(pinned.state.kind === 'granted' || pinned.state.kind === 'prompt') && (
          <PinnedFolderSection
            state={pinned.state}
            onReconnect={() => void pinned.reconnect()}
            onOpenFile={async (file) => {
              loading.set({ fileName: file.name, sizeBytes: file.size, phase: 'reading' });
              try {
                await loadSpreadsheetFile(file, null, wb.replaceWorkbook, (phase) =>
                  loading.set({ phase }),
                );
                onDismiss();
                loading.set(null);
              } catch (err) {
                console.error('[home] pinned-folder open failed', err);
                loading.set({
                  fileName: file.name,
                  phase: 'reading',
                  error: err instanceof Error ? err.message : 'Could not open this file.',
                });
              }
            }}
          />
        )}

        {recents.length > 0 && (
          <section className="home__section home__section--recents">
            <div className="home__section-head">
              <h2>
                <Icon name="history" /> Recent files
              </h2>
              <span className="home__section-hint">From your last sessions.</span>
            </div>
            <ul className="home__recents">
              {recents.map((rec) => (
                <li key={rec.id} className="home__recent">
                  <button
                    type="button"
                    className="home__recent-open"
                    onClick={() => onOpenRecent(rec)}
                    data-testid="home-recent-open"
                  >
                    <span className="home__recent-icon">
                      <Icon name="description" />
                    </span>
                    <span className="home__recent-text">
                      <span className="home__recent-name">{rec.name}</span>
                      <span className="home__recent-meta">
                        {formatSize(rec.size)} · {formatTime(rec.modifiedAt)}
                      </span>
                    </span>
                  </button>
                  <button
                    type="button"
                    className="home__recent-delete"
                    title="Remove from recent"
                    onClick={() => onDeleteRecent(rec)}
                  >
                    <Icon name="close" size="sm" />
                  </button>
                </li>
              ))}
            </ul>
          </section>
        )}

        <footer className="home__foot">
          <span>
            Drop an Excel / ODS / CSV file anywhere, or use <strong>File → Open</strong>.
          </span>
        </footer>
      </div>
    </div>
  );
}

function PinFolderControl({
  state,
  onPin,
  onReconnect,
  onUnpin,
}: {
  state: ReturnType<typeof usePinnedFolder>['state'];
  onPin: () => void;
  onReconnect: () => void;
  onUnpin: () => void;
}) {
  if (state.kind === 'unsupported') return null;
  if (state.kind === 'none' || state.kind === 'denied') {
    return (
      <button
        type="button"
        className="home__pin"
        onClick={onPin}
        data-testid="home-pin-folder"
        title={
          state.kind === 'denied'
            ? `Pick a folder again (last pick "${state.record.name}" was denied)`
            : 'Pick a folder — saves write directly to it, no download'
        }
      >
        <Icon name="folder_special" />
        <span>{state.kind === 'denied' ? `Re-pin ${state.record.name}` : 'Pin a folder'}</span>
      </button>
    );
  }
  if (state.kind === 'prompt') {
    return (
      <button
        type="button"
        className="home__pin home__pin--reconnect"
        onClick={onReconnect}
        data-testid="home-reconnect-folder"
        title="Permission lapsed — click to re-grant access"
      >
        <Icon name="link_off" />
        <span>Reconnect {state.record.name}</span>
      </button>
    );
  }
  // granted
  return (
    <div className="home__pin home__pin--granted" data-testid="home-pinned-folder">
      <Icon name="folder_special" />
      <span className="home__pin-name">Saving to {state.record.name}</span>
      <button
        type="button"
        className="home__pin-unpin"
        onClick={onUnpin}
        aria-label="Unpin folder"
        title="Unpin folder"
        data-testid="home-unpin-folder"
      >
        <Icon name="close" size="sm" />
      </button>
    </div>
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
