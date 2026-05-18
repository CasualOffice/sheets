import type { IWorkbookData } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';

/**
 * Workbook metadata kept under `IWorkbookData.custom.properties`.
 * Survives save/load round-trips because Univer preserves `custom`.
 */
export type WorkbookProperties = {
  title?: string;
  subject?: string;
  author?: string;
  tags?: string;
  category?: string;
  description?: string;
  /** Company / manager — surface as separate fields in xlsx App
   *  Properties (`docProps/app.xml`), used by Office's right-click
   *  details pane. Kept editable so a self-hosted deploy can preset
   *  per-org branding. */
  company?: string;
  manager?: string;
  /** ISO timestamps. */
  createdAt?: string;
  modifiedAt?: string;
};

const KEY = 'properties';

export function readProperties(api: FUniver): WorkbookProperties {
  const wb = api.getActiveWorkbook();
  if (!wb) return {};
  const snap = wb.save() as IWorkbookData;
  const props = (snap.custom?.[KEY] as WorkbookProperties | undefined) ?? {};
  return props;
}

/**
 * Write properties to the workbook's `custom.properties` slot.
 *
 * Univer doesn't expose a "patch custom field" command, so we read the
 * snapshot, mutate the field, and re-publish via the workbook's facade.
 * Because we're writing to `custom` (not cell data), no mutation events fire
 * — that's fine, the file metadata isn't part of cell-level state.
 */
export function writeProperties(api: FUniver, patch: WorkbookProperties) {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const underlying: any = wb.getWorkbook();
  if (typeof underlying.setCustomMetadata !== 'function') {
    // Univer 0.22 doesn't expose a public custom-metadata setter; mutate the
    // snapshot's `custom` field through the workbook's internal config.
    const snapshot = underlying.getSnapshot() as IWorkbookData;
    snapshot.custom = {
      ...snapshot.custom,
      [KEY]: { ...(snapshot.custom?.[KEY] as object), ...patch },
    };
    return;
  }
  const current = underlying.getCustomMetadata?.() ?? {};
  underlying.setCustomMetadata({
    ...current,
    [KEY]: { ...(current[KEY] ?? {}), ...patch },
  });
}

/**
 * Computed properties — these are always derived from the live workbook,
 * never persisted.
 */
export type ComputedProperties = {
  sheetCount: number;
  cellCount: number;
  snapshotBytes: number;
};

export function computeProperties(api: FUniver): ComputedProperties {
  const wb = api.getActiveWorkbook();
  if (!wb) return { sheetCount: 0, cellCount: 0, snapshotBytes: 0 };

  const snap = wb.save() as IWorkbookData;
  const sheetCount = snap.sheetOrder.length;

  let cellCount = 0;
  for (const id of snap.sheetOrder) {
    const ws = snap.sheets[id];
    if (!ws?.cellData) continue;
    const cd = ws.cellData as Record<string, Record<string, unknown>>;
    for (const r of Object.keys(cd)) {
      cellCount += Object.keys(cd[r] ?? {}).length;
    }
  }

  const snapshotBytes = new Blob([JSON.stringify(snap)]).size;
  return { sheetCount, cellCount, snapshotBytes };
}

export function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(2)} MB`;
}

export function formatDate(iso: string | undefined): string {
  if (!iso) return '—';
  try {
    return new Date(iso).toLocaleString();
  } catch {
    return '—';
  }
}
