import type { CSSProperties, ReactNode } from 'react';
import { TitleBar, type TitleBarProps } from './TitleBar';
import { Toolbar, type ToolbarProps } from './Toolbar';
import { FormulaBar, type FormulaBarProps } from './FormulaBar';
import { SheetTabs, type SheetTabsProps } from './SheetTabs';
import { StatusBar, type StatusBarProps } from './StatusBar';

export interface SheetShellProps {
  /** Chrome props — pass `null` to hide that strip. */
  titleBar?: TitleBarProps | null;
  toolbar?: ToolbarProps | null;
  formulaBar?: FormulaBarProps | null;
  sheetTabs?: SheetTabsProps | null;
  statusBar?: StatusBarProps | null;
  /** The canvas slot — typically `<CasualSheets />`. */
  children?: ReactNode;
  style?: CSSProperties;
}

/**
 * Cohesive editor shell — stacks TitleBar / Toolbar / FormulaBar / canvas /
 * SheetTabs / StatusBar with the pixel heights from the design system. Each
 * strip can be hidden by passing `null` to that prop slot. The `children`
 * slot is the canvas — the embed runtime mounts `<CasualSheets>` there;
 * Drive can mount its own composition.
 */
export function SheetShell({
  titleBar,
  toolbar,
  formulaBar,
  sheetTabs,
  statusBar,
  children,
  style,
}: SheetShellProps) {
  return (
    <div
      className="cs-sheet-shell"
      style={{
        display: 'flex',
        flexDirection: 'column',
        width: '100%',
        height: '100%',
        background: 'var(--color-bg)',
        color: 'var(--color-text)',
        fontFamily: 'var(--font-sans)',
        ...style,
      }}
    >
      {titleBar && <TitleBar {...titleBar} />}
      {toolbar && <Toolbar {...toolbar} />}
      {formulaBar && <FormulaBar {...formulaBar} />}
      <div
        className="cs-canvas-slot"
        style={{
          flex: '1 1 auto',
          minHeight: 0,
          position: 'relative',
          background: 'var(--color-bg)',
        }}
      >
        {children}
      </div>
      {sheetTabs && <SheetTabs {...sheetTabs} />}
      {statusBar && <StatusBar {...statusBar} />}
    </div>
  );
}
