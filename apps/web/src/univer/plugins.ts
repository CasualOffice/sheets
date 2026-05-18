import type { Univer } from '@univerjs/core';

import { UniverRenderEnginePlugin } from '@univerjs/engine-render';
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula';
import { UniverRPCMainThreadPlugin } from '@univerjs/rpc';
import { UniverUIPlugin } from '@univerjs/ui';
import { UniverDocsPlugin } from '@univerjs/docs';
import { UniverDocsUIPlugin } from '@univerjs/docs-ui';
import { UniverSheetsPlugin } from '@univerjs/sheets';
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui';
import { UniverSheetsFormulaPlugin, CalculationMode } from '@univerjs/sheets-formula';
import { UniverSheetsFormulaUIPlugin } from '@univerjs/sheets-formula-ui';
import { UniverSheetsNumfmtPlugin } from '@univerjs/sheets-numfmt';
import { UniverSheetsNumfmtUIPlugin } from '@univerjs/sheets-numfmt-ui';

/**
 * Register the EAGER plugin set. These are the plugins every workbook
 * needs to render and compute (render/formula/RPC/UI/docs/sheets/
 * sheets-ui/sheets-formula/numfmt). All the feature-specific plugins
 * — CF, DV, hyperlink, table, note, thread-comment, drawing, sort,
 * filter, find-replace — are deferred and loaded lazily by
 * `lazy-plugins.ts`, either when the user reaches for them or when
 * the snapshot's resources reference their data (eager-load-on-mount).
 *
 * Plugin order matters: render/formula engines before sheets,
 * sheets before sheets-ui, every base plugin before its `-ui`
 * counterpart. If you add an eager plugin, slot it next to its peers.
 */
export function registerPlugins(univer: Univer, container: HTMLElement): void {
  univer.registerPlugin(UniverRenderEnginePlugin);
  // Formula engine runs in a worker — main thread still loads the plugin
  // (for shared types, mutations) but skips the actual compute. The
  // worker ships the heavy `evaluate` path so paste / sort / fill on
  // large workbooks doesn't freeze the UI thread. See
  // apps/web/src/univer/formula-worker.ts for the worker side.
  univer.registerPlugin(UniverFormulaEnginePlugin, {
    notExecuteFormula: true,
  });
  const worker = new Worker(new URL('./formula-worker.ts', import.meta.url), {
    type: 'module',
    name: 'formula-worker',
  });
  univer.registerPlugin(UniverRPCMainThreadPlugin, { workerURL: worker });
  univer.registerPlugin(UniverUIPlugin, {
    container,
    header: false,
    toolbar: false,
    footer: false,
    contextMenu: true,
  });
  univer.registerPlugin(UniverDocsPlugin);
  univer.registerPlugin(UniverDocsUIPlugin);
  univer.registerPlugin(UniverSheetsPlugin, { notExecuteFormula: true });
  univer.registerPlugin(UniverSheetsUIPlugin);
  univer.registerPlugin(UniverSheetsFormulaPlugin, {
    notExecuteFormula: true,
    initialFormulaComputing: CalculationMode.NO_CALCULATION,
  });
  univer.registerPlugin(UniverSheetsFormulaUIPlugin);
  univer.registerPlugin(UniverSheetsNumfmtPlugin);
  univer.registerPlugin(UniverSheetsNumfmtUIPlugin);
}
