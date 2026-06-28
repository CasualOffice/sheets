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

import { useMemo, useState } from 'react';
import type { IWorkbookData } from '@univerjs/core';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { useLiveVersionList } from '../version-history/useLiveVersionList';
import {
  deleteVersion,
  readVersion,
  renameVersion,
  type VersionSnapshot,
} from '../version-history/store';
import { Icon } from './Icon';
import { HistoryPanel } from './HistoryPanel';

/**
 * Two-tab side panel:
 *
 *   - **Versions**: snapshot list grouped by day. Click a row to open
 *     a preview in the main grid; from there Restore or Cancel via
 *     the banner. Manual entries can be renamed or deleted; auto
 *     entries are read-only (they prune by retention).
 *   - **Activity**: per-mutation log surfaced by `<HistoryPanel />` —
 *     useful inside co-edit rooms, retained for parity with the
 *     pre-version-history UX.
 *
 * Industry references: Google Sheets "File > Version history" and
 * Excel Online "File > Info > Version History". Both show a flat
 * time-grouped list with preview-in-place and an explicit Restore.
 */

type Tab = 'versions' | 'activity';

export function VersionHistoryPanel() {
  const [tab, setTab] = useState<Tab>('versions');
  return (
    <aside className="side-panel side-panel--history" data-testid="version-history-panel" aria-label="Version history">
      <header className="side-panel__header">
        <Icon name="history" size="sm" />
        <h2 className="side-panel__title">History</h2>
      </header>
      <div className="version-history__tabs" role="tablist">
        <button
          type="button"
          role="tab"
          aria-selected={tab === 'versions'}
          data-testid="version-history-tab-versions"
          className={`version-history__tab${tab === 'versions' ? ' version-history__tab--active' : ''}`}
          onClick={() => setTab('versions')}
        >
          Versions
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={tab === 'activity'}
          data-testid="version-history-tab-activity"
          className={`version-history__tab${tab === 'activity' ? ' version-history__tab--active' : ''}`}
          onClick={() => setTab('activity')}
        >
          Activity
        </button>
      </div>
      {tab === 'versions' ? <VersionsTab /> : <ActivityTab />}
    </aside>
  );
}

function VersionsTab() {
  const list = useLiveVersionList();
  const groups = useMemo(() => groupByDay(list), [list]);
  const wb = useWorkbook();
  const api = useUniverAPI();

  const onPreview = async (v: VersionSnapshot) => {
    if (v.id == null) return;
    if (!api) return;
    // Already previewing this version → no-op so a stray double-click
    // doesn't drop the saved-state ref.
    if (wb.preview?.versionId === v.id) return;
    const wbInst = api.getActiveWorkbook();
    if (!wbInst) return;
    const liveData = wbInst.save() as unknown as IWorkbookData;
    // Read the full snapshot here rather than carrying it through
    // listVersions — keeps the list payload small for the panel render.
    const full = await readVersion(v.id);
    if (!full) return;
    wb.enterPreview(
      v.id,
      v.name,
      v.savedAt,
      full.data,
      (full.sourceFormat as Parameters<typeof wb.enterPreview>[4]) ?? null,
      liveData,
      wb.meta.sourceFormat,
    );
  };

  if (list.length === 0) {
    return (
      <div className="side-panel__empty">
        No saved versions yet. Snapshots are captured automatically every
        ~10 minutes while you edit, or any time via File → Save version.
      </div>
    );
  }

  return (
    <div className="version-history__body" data-testid="version-history-versions">
      {groups.map((g) => (
        <div key={g.label} className="version-history__group">
          <h3 className="version-history__group-label">{g.label}</h3>
          <ul className="version-history__list" role="list">
            {g.items.map((v) => (
              <VersionRow
                key={v.id}
                v={v}
                isActive={wb.preview?.versionId === v.id}
                onPreview={() => void onPreview(v)}
              />
            ))}
          </ul>
        </div>
      ))}
    </div>
  );
}

function VersionRow({
  v,
  isActive,
  onPreview,
}: {
  v: VersionSnapshot;
  isActive: boolean;
  onPreview: () => void;
}) {
  const wb = useWorkbook();
  const [renaming, setRenaming] = useState(false);
  const [draftName, setDraftName] = useState(v.name);

  const onRename = async () => {
    if (v.id == null) return;
    const next = draftName.trim();
    if (!next || next === v.name) {
      setRenaming(false);
      return;
    }
    await renameVersion(v.id, next);
    setRenaming(false);
  };

  const onDelete = async () => {
    if (v.id == null) return;
    // If the row being deleted is the one currently previewed, exit
    // preview first so the saved-state ref restores cleanly.
    if (wb.preview?.versionId === v.id) wb.exitPreview();
    await deleteVersion(v.id);
  };

  return (
    <li
      className={`version-history__row${isActive ? ' version-history__row--active' : ''}`}
      data-testid="version-history-row"
    >
      <button
        type="button"
        className="version-history__row-main"
        data-testid="version-history-row-preview"
        onClick={onPreview}
        aria-label={`Preview ${v.name}`}
      >
        <span className={`version-history__kind version-history__kind--${v.kind}`} title={v.kind === 'manual' ? 'Saved by you' : 'Auto-saved'}>
          <Icon name={v.kind === 'manual' ? 'bookmark' : 'schedule'} size="sm" />
        </span>
        <span className="version-history__row-text">
          {renaming ? (
            <input
              autoFocus
              className="version-history__rename-input"
              value={draftName}
              onChange={(e) => setDraftName(e.target.value)}
              onBlur={() => void onRename()}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  void onRename();
                } else if (e.key === 'Escape') {
                  e.preventDefault();
                  setDraftName(v.name);
                  setRenaming(false);
                }
              }}
              onClick={(e) => e.stopPropagation()}
            />
          ) : (
            <span className="version-history__name">{v.name}</span>
          )}
          <span className="version-history__when">
            {formatTime(v.savedAt)}
            {v.size ? ` · ${formatSize(v.size)}` : ''}
          </span>
        </span>
      </button>
      <div className="version-history__row-actions">
        {v.kind === 'manual' && (
          <button
            type="button"
            className="version-history__action"
            data-testid="version-history-rename"
            title="Rename"
            onClick={(e) => {
              e.stopPropagation();
              setDraftName(v.name);
              setRenaming(true);
            }}
          >
            <Icon name="edit" size="sm" />
          </button>
        )}
        <button
          type="button"
          className="version-history__action version-history__action--danger"
          data-testid="version-history-delete"
          title="Delete this version"
          onClick={(e) => {
            e.stopPropagation();
            void onDelete();
          }}
        >
          <Icon name="delete" size="sm" />
        </button>
      </div>
    </li>
  );
}

function ActivityTab() {
  // Reuses the existing per-mutation panel body verbatim. HistoryPanel
  // renders its own outer aside; we wrap with a class that strips the
  // chrome we already provide here so the nesting doesn't double up.
  return (
    <div className="version-history__activity">
      <HistoryPanel />
    </div>
  );
}

type Group = { label: string; items: VersionSnapshot[] };

function groupByDay(list: VersionSnapshot[]): Group[] {
  const now = Date.now();
  const out: Group[] = [];
  const findOrCreate = (label: string): Group => {
    const found = out.find((g) => g.label === label);
    if (found) return found;
    const g: Group = { label, items: [] };
    out.push(g);
    return g;
  };
  for (const v of list) {
    const age = now - v.savedAt;
    let label: string;
    if (age < 24 * 60 * 60 * 1000 && isSameDay(v.savedAt, now)) label = 'Today';
    else if (isSameDay(v.savedAt, now - 24 * 60 * 60 * 1000)) label = 'Yesterday';
    else if (age < 7 * 24 * 60 * 60 * 1000) label = 'Earlier this week';
    else if (age < 30 * 24 * 60 * 60 * 1000) label = 'Earlier this month';
    else label = 'Older';
    findOrCreate(label).items.push(v);
  }
  return out;
}

function isSameDay(a: number, b: number): boolean {
  const da = new Date(a);
  const db = new Date(b);
  return (
    da.getFullYear() === db.getFullYear() &&
    da.getMonth() === db.getMonth() &&
    da.getDate() === db.getDate()
  );
}

function formatTime(ts: number): string {
  const d = new Date(ts);
  const now = Date.now();
  const delta = now - ts;
  if (delta < 60_000) return 'just now';
  if (delta < 60 * 60_000) return `${Math.floor(delta / 60_000)} min ago`;
  return d.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}
