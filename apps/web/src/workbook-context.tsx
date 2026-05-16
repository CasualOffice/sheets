import { createContext } from 'react';
import type { IWorkbookData } from '@univerjs/core';

/**
 * Lifts the active workbook snapshot to App state so File → Open can replace
 * it. Changing `snapshot` re-mounts `<UniverSheet>` (which keys on snapshot
 * identity), dropping the old Univer instance and creating a new one with
 * the loaded data.
 */
export type WorkbookCtxValue = {
  snapshot: IWorkbookData;
  replaceWorkbook: (next: IWorkbookData) => void;
};

export const WorkbookContext = createContext<WorkbookCtxValue | null>(null);
