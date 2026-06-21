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

/** Drop blank / placeholder string values so the dialog never renders
 *  literal `"null"`/`"undefined"`/`"Unknown"` left behind by some xlsx
 *  authoring tools. Mirrors the parser's `clean()` for older snapshots
 *  that were imported before that guard existed. */
export function sanitizeProps(raw: WorkbookProperties): WorkbookProperties {
  const out: WorkbookProperties = {};
  for (const [k, v] of Object.entries(raw)) {
    if (typeof v !== 'string') continue;
    const t = v.trim();
    if (!t) continue;
    const lower = t.toLowerCase();
    if (lower === 'unknown' || lower === 'null' || lower === 'undefined') continue;
    out[k as keyof WorkbookProperties] = t;
  }
  return out;
}

export function readProperties(api: FUniver): WorkbookProperties {
  const wb = api.getActiveWorkbook();
  if (!wb) return {};
  const snap = wb.save() as IWorkbookData;
  const props = (snap.custom?.[KEY] as WorkbookProperties | undefined) ?? {};
  return sanitizeProps(props);
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
  /** Workbook name — the uploaded filename (sans extension) for imported
   *  files, set in `openSpreadsheetFile`. */
  name: string;
  sheetCount: number;
  cellCount: number;
  /** Best-available byte size: the real on-disk size of the uploaded file
   *  when known (`custom.sourceBytes`), else an estimate from the in-memory
   *  JSON snapshot. */
  sizeBytes: number;
  /** True when `sizeBytes` is the actual uploaded file size; false when it's
   *  the (much larger, uncompressed) snapshot estimate. */
  sizeIsExact: boolean;
};

export function computeProperties(api: FUniver): ComputedProperties {
  const wb = api.getActiveWorkbook();
  if (!wb) return { name: '', sheetCount: 0, cellCount: 0, sizeBytes: 0, sizeIsExact: false };

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

  // Prefer the real uploaded size; the JSON snapshot is uncompressed and runs
  // several times larger than the zipped xlsx, so it's only a rough fallback
  // for workbooks created in-app (which were never a file on disk).
  const sourceBytes = snap.custom?.sourceBytes;
  const sizeIsExact = typeof sourceBytes === 'number' && sourceBytes > 0;
  const sizeBytes = sizeIsExact ? (sourceBytes as number) : new Blob([JSON.stringify(snap)]).size;

  return { name: snap.name ?? '', sheetCount, cellCount, sizeBytes, sizeIsExact };
}

export function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(2)} MB`;
}

export function formatDate(iso: string | undefined): string {
  if (!iso) return '—';
  // `new Date('garbage')` returns an Invalid Date object rather than throwing,
  // and its toLocaleString() is the literal "Invalid Date" — guard for it so a
  // malformed timestamp shows an em-dash instead of garbage.
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '—';
  return d.toLocaleString();
}
