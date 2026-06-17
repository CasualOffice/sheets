/**
 * @schnsrw/casual-sheets/shell — editor chrome lifted from the design
 * bundle. Props-driven; consumers compose these with the canvas they need
 * (typically `<CasualSheets>` from the same package).
 *
 * The shell components consume `@schnsrw/design-system` for tokens +
 * primitives, so consumers must import `@schnsrw/design-system/tokens.css`
 * at app entry once.
 */

export { TitleBar } from './TitleBar';
export type { MenuDescriptor, TitleBarProps } from './TitleBar';

export { Toolbar } from './Toolbar';
export type { ToolbarCallbacks, ToolbarFormatState, ToolbarProps } from './Toolbar';

export { FormulaBar } from './FormulaBar';
export type { FormulaBarProps } from './FormulaBar';

export { SheetTabs } from './SheetTabs';
export type { SheetDescriptor, SheetTabsProps } from './SheetTabs';

export { StatusBar } from './StatusBar';
export type { SelectionStats, StatusBarProps } from './StatusBar';

export { SheetShell } from './SheetShell';
export type { SheetShellProps } from './SheetShell';
