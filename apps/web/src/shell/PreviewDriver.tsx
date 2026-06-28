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

import { useEffect } from 'react';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { applyViewOnlyMode } from '../collab/view-mode';

/**
 * Enforces read-only while a version-history preview is active. The
 * existing `applyViewOnlyMode` helper (built for view-only co-edit
 * joiners) flips Univer's `WorkbookEditablePermission` to `false`,
 * which makes the cell editor refuse to open and disables mutating
 * menu items. Returns a restorer; we hold it for the lifetime of the
 * preview and call it on exit.
 *
 * Watches `meta.revision` so the re-apply also fires when
 * `enterPreview` bumps the revision (the workbook swap recreates the
 * permission point, so we have to set it again after the swap).
 */
export function PreviewDriver() {
  const api = useUniverAPI();
  const { preview, meta } = useWorkbook();

  useEffect(() => {
    if (!api || !preview) return;
    const wb = api.getActiveWorkbook();
    if (!wb) return;
    const unitId = wb.getId();
    // Univer 0.22's permission point isn't always registered the
    // microtask after a workbook swap — wait one rAF before applying.
    let restorer: (() => void) | null = null;
    const raf = requestAnimationFrame(() => {
      restorer = applyViewOnlyMode(api, unitId);
    });
    return () => {
      cancelAnimationFrame(raf);
      try {
        restorer?.();
      } catch (err) {
        console.warn('[preview-driver] restore failed', err);
      }
    };
    // meta.revision included so we re-apply after a re-swap inside
    // the preview (rare, but the workbook id may change).
  }, [api, preview, meta.revision]);

  return null;
}
