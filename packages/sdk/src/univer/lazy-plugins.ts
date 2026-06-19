import type { Univer, IWorkbookData } from '@univerjs/core';

/**
 * Lazy plugin loading — pipeline Stage 4. The cheap plugins (render,
 * formula, sheets, sheets-ui, numfmt, docs) stay eager because every
 * workbook needs them. The heavy / feature-specific plugins (CF, DV,
 * hyperlink, table, note, thread-comment, drawing, sort, filter,
 * find-replace) load on demand:
 *
 *   1. Eagerly when a snapshot is about to mount and we detect the
 *      plugin's resource key on it (e.g. `SHEET_CONDITIONAL_FORMATTING_PLUGIN`
 *      in `data.resources`). This is the safety net: missing this would
 *      silently drop plugin data on file open.
 *   2. Lazily when the user reaches for the feature — opens the Data
 *      tab (sort/filter), hits Ctrl+F (find-replace), uses the Insert
 *      tab (drawing), etc. The shell hooks `ensurePlugin(...)` into
 *      these triggers and awaits before the action runs.
 *
 * Each loader returns the plugins-to-register in the correct order
 * (base before UI, same as `plugins.ts`). Registration order matters
 * in Univer; the loaders bundle a small group together to keep that
 * locality explicit.
 */

export type LazyPluginGroup =
  | 'cf'
  | 'dv'
  | 'hyperlink'
  | 'note'
  | 'table'
  | 'threadComment'
  | 'drawing'
  | 'sort'
  | 'filter'
  | 'findReplace';

type Loader = () => Promise<Array<[unknown, unknown?]>>;

const LOADERS: Record<LazyPluginGroup, Loader> = {
  cf: async () => {
    const [base, ui] = await Promise.all([
      import('@univerjs/sheets-conditional-formatting'),
      import('@univerjs/sheets-conditional-formatting-ui'),
    ]);
    return [
      [base.UniverSheetsConditionalFormattingPlugin],
      [ui.UniverSheetsConditionalFormattingUIPlugin],
    ];
  },
  dv: async () => {
    const [base, ui] = await Promise.all([
      import('@univerjs/sheets-data-validation'),
      import('@univerjs/sheets-data-validation-ui'),
    ]);
    return [
      [base.UniverSheetsDataValidationPlugin],
      [ui.UniverSheetsDataValidationUIPlugin],
    ];
  },
  hyperlink: async () => {
    const [base, ui] = await Promise.all([
      import('@univerjs/sheets-hyper-link'),
      import('@univerjs/sheets-hyper-link-ui'),
    ]);
    return [
      [base.UniverSheetsHyperLinkPlugin],
      [ui.UniverSheetsHyperLinkUIPlugin],
    ];
  },
  note: async () => {
    const [base, ui] = await Promise.all([
      import('@univerjs/sheets-note'),
      import('@univerjs/sheets-note-ui'),
    ]);
    return [[base.UniverSheetsNotePlugin], [ui.UniverSheetsNoteUIPlugin]];
  },
  table: async () => {
    const [base, ui] = await Promise.all([
      import('@univerjs/sheets-table'),
      import('@univerjs/sheets-table-ui'),
    ]);
    return [[base.UniverSheetsTablePlugin], [ui.UniverSheetsTableUIPlugin]];
  },
  threadComment: async () => {
    const [tc, tcUi, sheetsTc, sheetsTcUi] = await Promise.all([
      import('@univerjs/thread-comment'),
      import('@univerjs/thread-comment-ui'),
      import('@univerjs/sheets-thread-comment'),
      import('@univerjs/sheets-thread-comment-ui'),
    ]);
    return [
      [tc.UniverThreadCommentPlugin],
      [tcUi.UniverThreadCommentUIPlugin],
      [sheetsTc.UniverSheetsThreadCommentPlugin],
      [sheetsTcUi.UniverSheetsThreadCommentUIPlugin],
    ];
  },
  drawing: async () => {
    const [d, dUi, sd, sdUi] = await Promise.all([
      import('@univerjs/drawing'),
      import('@univerjs/drawing-ui'),
      import('@univerjs/sheets-drawing'),
      import('@univerjs/sheets-drawing-ui'),
      // Side-effect imports: install FWorksheet.insertImage / getImages /
      // updateImages on the facade prototype. Without these, code that
      // reaches in via the FUniver facade (e2e specs, future shell glue)
      // sees an undefined method even though the plugin is registered.
      import('@univerjs/sheets-drawing/facade'),
      import('@univerjs/sheets-drawing-ui/facade'),
    ]);
    return [
      [d.UniverDrawingPlugin],
      [dUi.UniverDrawingUIPlugin],
      [sd.UniverSheetsDrawingPlugin],
      [sdUi.UniverSheetsDrawingUIPlugin],
    ];
  },
  sort: async () => {
    const [base, ui] = await Promise.all([
      import('@univerjs/sheets-sort'),
      import('@univerjs/sheets-sort-ui'),
    ]);
    return [[base.UniverSheetsSortPlugin], [ui.UniverSheetsSortUIPlugin]];
  },
  filter: async () => {
    const [base, ui] = await Promise.all([
      import('@univerjs/sheets-filter'),
      import('@univerjs/sheets-filter-ui'),
    ]);
    return [[base.UniverSheetsFilterPlugin], [ui.UniverSheetsFilterUIPlugin]];
  },
  findReplace: async () => {
    const [base, sheets] = await Promise.all([
      import('@univerjs/find-replace'),
      import('@univerjs/sheets-find-replace'),
    ]);
    return [[base.UniverFindReplacePlugin], [sheets.UniverSheetsFindReplacePlugin]];
  },
};

/**
 * Map from a Univer `data.resources[].name` to the lazy group that owns
 * it. Used by `eagerLoadForSnapshot` to pre-register plugins whose
 * state already lives on the workbook so file-open never silently
 * drops a CF rule / table / drawing / etc.
 */
const RESOURCE_NAME_TO_GROUP: Record<string, LazyPluginGroup> = {
  SHEET_CONDITIONAL_FORMATTING_PLUGIN: 'cf',
  SHEET_DATA_VALIDATION_PLUGIN: 'dv',
  SHEET_HYPER_LINK_PLUGIN: 'hyperlink',
  SHEET_NOTE_PLUGIN: 'note',
  SHEET_TABLE_PLUGIN: 'table',
  SHEET_THREAD_COMMENT_BASE_PLUGIN: 'threadComment',
  SHEET_DRAWING_PLUGIN: 'drawing',
  SHEET_SORT_PLUGIN: 'sort',
  SHEET_FILTER_PLUGIN: 'filter',
};

const loaded = new Set<LazyPluginGroup>();
const inflight = new Map<LazyPluginGroup, Promise<void>>();

/**
 * Module-level reference to the live Univer instance, set by
 * `UniverSheet.tsx` immediately after `new Univer()`. Lets shell code
 * call `ensurePluginByName(group)` without plumbing the Univer
 * instance through React context (the FUniver facade doesn't expose
 * its host). Cleared on dispose so callers fail loudly if they hit
 * the lazy path after teardown.
 */
let currentUniver: Univer | null = null;

export function setUniverForLazyLoad(univer: Univer | null): void {
  currentUniver = univer;
}

/**
 * Idempotent: subsequent calls for the same group resolve immediately
 * if the plugin is already loaded; concurrent calls share the same
 * in-flight promise so we never double-register.
 */
export function ensurePlugin(univer: Univer, group: LazyPluginGroup): Promise<void> {
  if (loaded.has(group)) return Promise.resolve();
  const existing = inflight.get(group);
  if (existing) return existing;
  const loader = LOADERS[group];
  if (!loader) return Promise.resolve();
  const p = loader().then((plugins) => {
    for (const [PluginCtor, config] of plugins) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      univer.registerPlugin(PluginCtor as any, config);
    }
    loaded.add(group);
    inflight.delete(group);
  });
  inflight.set(group, p);
  return p;
}

/**
 * Shell-friendly variant of `ensurePlugin` that pulls the Univer
 * instance from the module-level holder set by `UniverSheet.tsx`. Use
 * this anywhere we only have the FUniver facade (toolbar callbacks,
 * panel handlers). Returns a resolved promise if the holder is empty —
 * the worst case is a no-op, never a throw.
 */
export function ensurePluginByName(group: LazyPluginGroup): Promise<void> {
  if (loaded.has(group)) return Promise.resolve();
  if (!currentUniver) return Promise.resolve();
  return ensurePlugin(currentUniver, group);
}

/**
 * Walk a snapshot for plugin-owned resources + side-channel hyperlinks
 * and eagerly load every group that's referenced. Returns a promise
 * the caller MUST await before `createUnit` — Univer's resource manager
 * silently discards keys for plugins that aren't yet registered.
 */
export async function eagerLoadForSnapshot(univer: Univer, snapshot: IWorkbookData): Promise<void> {
  const groups = new Set<LazyPluginGroup>();
  const resources = snapshot.resources ?? [];
  for (const r of resources) {
    const g = RESOURCE_NAME_TO_GROUP[r.name];
    if (g) groups.add(g);
  }
  // Hyperlinks live inline in cell.p (Stage 5) — the hyperlink plugin
  // still owns click handling / context-menu actions, so eager-load
  // when any cell has a HYPERLINK customRange.
  if (snapshotHasHyperlinks(snapshot)) groups.add('hyperlink');
  await Promise.all(Array.from(groups).map((g) => ensurePlugin(univer, g)));
}

/**
 * Idle-load every remaining lazy group AFTER Univer is mounted. The
 * bundle split (each group ships as its own chunk) is the persistent
 * boot-time win — the initial paint doesn't pay for them. This
 * idle-load just ensures they're all eventually registered so a user
 * clicking "Insert > Table" doesn't hit a no-op.
 *
 * Loads in parallel via `requestIdleCallback` so they don't compete
 * with the first paint, falling back to `setTimeout(0)` in browsers
 * without rIC (Safari).
 */
export function idleLoadAll(univer: Univer): void {
  const groups = Object.keys(LOADERS) as LazyPluginGroup[];
  schedule(() => {
    for (const g of groups) {
      void ensurePlugin(univer, g);
    }
  });
}

function schedule(fn: () => void): void {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ric = (globalThis as any).requestIdleCallback as
    | ((cb: () => void, opts?: { timeout: number }) => number)
    | undefined;
  if (ric) ric(fn, { timeout: 500 });
  else setTimeout(fn, 0);
}

function snapshotHasHyperlinks(snapshot: IWorkbookData): boolean {
  const sheetOrder = snapshot.sheetOrder ?? [];
  for (const sid of sheetOrder) {
    const sheet = snapshot.sheets?.[sid];
    if (!sheet?.cellData) continue;
    const cellData = sheet.cellData as Record<
      string,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      Record<string, { p?: any }>
    >;
    for (const r of Object.keys(cellData)) {
      const row = cellData[r];
      for (const c of Object.keys(row)) {
        const ranges = row[c]?.p?.body?.customRanges ?? [];
        if (ranges.some((cr: { rangeType?: number }) => cr.rangeType === 0)) return true;
      }
    }
  }
  return false;
}
