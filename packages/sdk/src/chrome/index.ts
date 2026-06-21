/**
 * Chrome — the Office shell as slot components for `<CasualSheets chrome>`.
 *
 * First slice (SDK_MIGRATION_PIPELINE Phase 1 step 2): a minimal built-in
 * `Toolbar`. The rich shell (FormulaBar, MenuBar, TitleBar, StatusBar) lifts
 * from `apps/web/src/shell/` here in later slices, behind `chrome="full"`.
 */
export { Toolbar, type ToolbarProps } from './Toolbar';
export { FormulaBar, type FormulaBarProps } from './FormulaBar';
export { StatusBar, type StatusBarProps } from './StatusBar';
export { Icon, type IconProps } from './Icon';
export { ensureChromeFonts } from './fonts';
export { ColorPicker, type ColorPickerProps } from './ColorPicker';
export { BordersPicker, type BordersPickerProps } from './BordersPicker';
export { AutoSumPicker, type AutoSumPickerProps } from './AutoSumPicker';
export { MenuBar, type MenuBarProps, type MenuDialogKind } from './MenuBar';
export { NameBox, type NameBoxProps } from './NameBox';
export { SheetTabs, type SheetTabsProps } from './SheetTabs';
export { FindReplace, type FindReplaceProps } from './FindReplace';
// Dialog foundation: the modal primitive, the dialog host/registry, the
// extension API, and the built-in reference dialogs.
export { Dialog, type DialogProps } from './Dialog';
export { FormatCellsDialog } from './FormatCellsDialog';
export {
  DialogProvider,
  useDialogs,
  hasBuiltInDialog,
  type DialogKind,
  type DialogController,
  type DialogProviderProps,
} from './dialog-context';
export type {
  ChromeExtensions,
  ToolbarExtension,
  MenuExtension,
  PanelExtension,
  DialogExtension,
  DialogComponentProps,
  PanelComponentProps,
  MenuTarget,
} from './extensions';
// Lazy-loaded shells: chrome above (ChromeTop) and below (ChromeBottom) the grid,
// imported on demand by `<CasualSheets>` via the `@casualoffice/sheets/chrome`
// subpath so `chrome="none"` consumers don't bundle the chrome JS.
export { ChromeTop } from './ChromeTop';
export { ChromeBottom } from './ChromeBottom';
