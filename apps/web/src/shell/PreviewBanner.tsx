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

import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { saveNamedVersion } from '../version-history/useVersionHistoryCapture';
import { Icon } from './Icon';
import type { IWorkbookData } from '@univerjs/core';

/**
 * Banner pinned across the top of the grid host while a version-history
 * preview is active. Industry-standard layout (Google Sheets / Excel
 * Online): "Previewing version <name> · saved <relative time>" on the
 * left, Restore + Cancel buttons on the right.
 *
 * Hidden when `wb.preview` is null. Renders inline in the App layout,
 * outside the grid host, so it doesn't conflict with Univer's own
 * canvas events.
 */
export function PreviewBanner() {
  const wb = useWorkbook();
  const api = useUniverAPI();
  if (!wb.preview) return null;

  const onRestore = async () => {
    if (!wb.preview) return;
    // Capture the *currently visible* state (which IS the snapshot,
    // since we already swapped to it in enterPreview) as a manual
    // version named "Restored: …". That gives an explicit undo entry
    // in the history list. Skip if api isn't ready — the user can
    // still proceed without the undo crumb.
    if (api) {
      try {
        const data = api.getActiveWorkbook()?.save() as unknown as IWorkbookData | undefined;
        if (data) {
          await saveNamedVersion(
            data,
            `Restored ${wb.preview.versionName}`,
            wb.meta.sourceFormat ?? null,
          );
        }
      } catch (err) {
        console.warn('[preview] could not capture restore crumb', err);
      }
    }
    wb.commitPreview();
  };

  return (
    <div className="preview-banner" data-testid="preview-banner" role="status">
      <Icon name="history_toggle_off" size="sm" />
      <span className="preview-banner__text">
        Previewing <strong>{wb.preview.versionName}</strong> · saved {timeAgo(wb.preview.versionSavedAt)}. Editing is disabled until you restore or cancel.
      </span>
      <button
        type="button"
        className="preview-banner__btn"
        data-testid="preview-banner-cancel"
        onClick={wb.exitPreview}
      >
        Cancel
      </button>
      <button
        type="button"
        className="preview-banner__btn preview-banner__btn--primary"
        data-testid="preview-banner-restore"
        onClick={() => void onRestore()}
      >
        Restore this version
      </button>
    </div>
  );
}

function timeAgo(ts: number): string {
  const ms = Date.now() - ts;
  if (ms < 60_000) return 'moments ago';
  const mins = Math.round(ms / 60_000);
  if (mins < 60) return `${mins} min ago`;
  const hours = Math.round(mins / 60);
  if (hours < 24) return `${hours} hr ago`;
  const days = Math.round(hours / 24);
  return `${days} day${days === 1 ? '' : 's'} ago`;
}
