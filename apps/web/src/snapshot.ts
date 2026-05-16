import { LocaleType, type IWorkbookData } from '@univerjs/core';

/**
 * Initial workbook size. Univer materializes row/column metadata for the
 * declared count, so we keep this modest to boot fast. The grid grows
 * dynamically (see `useWorkbookGrowth`) up to MAX_ROWS / MAX_COLUMNS.
 */
export const INITIAL_ROWS = 1024;
export const INITIAL_COLUMNS = 128;
export const MAX_ROWS = 8192;
export const MAX_COLUMNS = 1024;

export function emptyWorkbook(): IWorkbookData {
  return {
    id: 'workbook-1',
    rev: 1,
    name: 'Untitled',
    appVersion: '0.22.1',
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
