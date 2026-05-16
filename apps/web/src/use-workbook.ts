import { useContext } from 'react';
import { WorkbookContext, type WorkbookCtxValue } from './workbook-context';

export function useWorkbook(): WorkbookCtxValue {
  const ctx = useContext(WorkbookContext);
  if (!ctx) throw new Error('useWorkbook must be used inside <WorkbookContext.Provider>');
  return ctx;
}
