/**
 * Snapshot defaults pulled out of apps/web/src/snapshot.ts so the
 * SDK's xlsx parser doesn't have to import that whole module. Keep
 * these in sync with the host snapshot module if the defaults drift.
 *
 * UNIVER_VERSION must match the runtime Univer the host boots — the
 * appVersion field on the IWorkbookData snapshot is checked at unit
 * mount and a mismatch warns in dev. Sheet apps' `../snapshot` reads
 * the version from the workspace's @univerjs/core dep; we hardcode
 * the same minor here because the SDK declares @univerjs/* as
 * `^0.24.0` peer.
 */
export const INITIAL_ROWS = 1024;
export const INITIAL_COLUMNS = 26;
export const UNIVER_VERSION = '0.24.0';
