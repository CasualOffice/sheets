/**
 * Formula calculation worker. Runs a stripped-down Univer in a Web Worker
 * so the main thread doesn't hang on large-workbook recalcs. The main
 * thread instance owns the data model + UI; this one only computes.
 *
 * Communication is via Univer's RPC plugins — UniverRPCMainThreadPlugin
 * on the main thread (wires the Worker) and UniverRPCWorkerThreadPlugin
 * here. Mutations flow over postMessage transparently to our code.
 *
 * The worker mirrors a minimal plugin set: SheetsPlugin (data model,
 * mutation-only mode), FormulaEnginePlugin (the actual calculator),
 * RemoteSheetsFormulaPlugin (the cross-boundary glue), FilterPlugin
 * (filter-aware formula evaluation, see Univer's reference worker).
 */
import { LocaleType, LogLevel, Univer } from '@univerjs/core';
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula';
import { UniverRPCWorkerThreadPlugin } from '@univerjs/rpc';
import { UniverSheetsPlugin } from '@univerjs/sheets';
import { UniverSheetsFilterPlugin } from '@univerjs/sheets-filter';
import { UniverRemoteSheetsFormulaPlugin } from '@univerjs/sheets-formula';

const univer = new Univer({
  locale: LocaleType.EN_US,
  logLevel: LogLevel.WARN,
});

univer.registerPlugins([
  // `onlyRegisterFormulaRelatedMutations: true` skips the half of the
  // sheets plugin that mutates UI / selection — the worker doesn't
  // render anything, it just consumes data mutations to compute formulas.
  [UniverSheetsPlugin, { onlyRegisterFormulaRelatedMutations: true }],
  [UniverFormulaEnginePlugin],
  [UniverRPCWorkerThreadPlugin],
  [UniverRemoteSheetsFormulaPlugin],
  // Filters affect formula visibility (SUBTOTAL etc.); keep the plugin
  // in the worker so calculations see the same hidden state as the UI.
  [UniverSheetsFilterPlugin],
]);

// Expose for debugging from the worker's devtools. WorkerGlobalScope isn't
// in our tsconfig lib (we only ship DOM types — adding WebWorker globally
// pollutes the main bundle), so use a loose `globalThis` cast instead.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).univer = univer;
