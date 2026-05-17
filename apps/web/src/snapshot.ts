import { LocaleType, type IWorkbookData } from '@univerjs/core';
import appPkg from '../package.json';

// Sourced from our package.json's `@univerjs/core` dependency declaration so
// the snapshot's appVersion tracks the installed Univer release automatically.
// (Vite supports JSON imports; tsconfig has resolveJsonModule.)
const UNIVER_DEP = (appPkg.dependencies as Record<string, string>)['@univerjs/core'];
export const UNIVER_VERSION = UNIVER_DEP.replace(/^[~^]/, '');

/**
 * Initial workbook size. Univer materializes row/column metadata for the
 * declared count, so we keep this modest to boot fast. The grid grows
 * dynamically (see `useWorkbookGrowth`) up to MAX_ROWS / MAX_COLUMNS.
 */
export const INITIAL_ROWS = 1024;
// 26 = A..Z. Univer allocates row/column metadata up-front for the declared
// count, so a 128-wide start cost ~5× the boot allocation for columns the
// user almost never reaches before `useWorkbookGrowth` extends them.
export const INITIAL_COLUMNS = 26;
export const MAX_ROWS = 8192;
export const MAX_COLUMNS = 1024;

export function emptyWorkbook(): IWorkbookData {
  return {
    // Unique per call — Univer's IUniverInstanceService rejects duplicate unit
    // ids, so a fresh blank workbook must not collide with the one it's
    // replacing.
    id: `wb-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
    rev: 1,
    name: 'Untitled',
    appVersion: UNIVER_VERSION,
    locale: LocaleType.EN_US,
    styles: {},
    sheetOrder: ['sheet-1'],
    sheets: {
      'sheet-1': {
        id: 'sheet-1',
        name: 'Sheet1',
        cellData: {},
        rowCount: INITIAL_ROWS,
        columnCount: INITIAL_COLUMNS,
      },
    },
  };
}
