/**
 * Chrome — the Office shell as slot components for `<CasualSheets chrome>`.
 *
 * First slice (SDK_MIGRATION_PIPELINE Phase 1 step 2): a minimal built-in
 * `Toolbar`. The rich shell (FormulaBar, MenuBar, TitleBar, StatusBar) lifts
 * from `apps/web/src/shell/` here in later slices, behind `chrome="full"`.
 */
export { Toolbar, type ToolbarProps } from './Toolbar';
export { FormulaBar, type FormulaBarProps } from './FormulaBar';
