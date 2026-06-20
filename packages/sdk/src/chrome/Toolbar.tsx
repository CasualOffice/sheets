/**
 * Toolbar — the minimal built-in Office chrome for `<CasualSheets chrome>`.
 *
 * This is the FIRST slice of the chrome lift (SDK_MIGRATION_PIPELINE Phase 1
 * step 2). It is intentionally small and self-contained: it drives the editor
 * purely through `CasualSheetsAPI.executeCommand` (no app context, no FileSource,
 * no design-system font dependency), so it works in any host out of the box.
 *
 * The rich Office shell from `apps/web/src/shell/` (FormulaBar, MenuBar, …) is
 * lifted in later slices behind `chrome="full"`, converging with the design
 * system. Until then `chrome="minimal"` and `chrome="full"` both render this.
 */

import type { CSSProperties } from 'react';
import type { CasualSheetsAPI } from '../sheets/api';

interface ToolbarAction {
  id: string;
  label: string;
  /** Univer command id dispatched via the facade. */
  command: string;
  glyph: string;
}

// Univer OSS command ids (same set the embed-runtime toolbar bridge maps).
const ACTIONS: ToolbarAction[] = [
  { id: 'undo', label: 'Undo', command: 'univer.command.undo', glyph: '↶' },
  { id: 'redo', label: 'Redo', command: 'univer.command.redo', glyph: '↷' },
  { id: 'bold', label: 'Bold', command: 'sheet.command.set-range-bold', glyph: 'B' },
  { id: 'italic', label: 'Italic', command: 'sheet.command.set-range-italic', glyph: 'I' },
  {
    id: 'underline',
    label: 'Underline',
    command: 'sheet.command.set-range-underline',
    glyph: 'U',
  },
];

const BAR_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 2,
  padding: '4px 6px',
  borderBottom: '1px solid var(--cs-chrome-border, rgba(0,0,0,0.12))',
  background: 'var(--cs-chrome-bg, #f8f9fa)',
  flex: '0 0 auto',
  userSelect: 'none',
};

const BTN_STYLE: CSSProperties = {
  minWidth: 28,
  height: 28,
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  border: 'none',
  borderRadius: 4,
  background: 'transparent',
  cursor: 'pointer',
  font: 'inherit',
  fontSize: 14,
  lineHeight: 1,
  color: 'var(--cs-chrome-fg, #1f2329)',
};

const ITALIC_GLYPH: CSSProperties = { fontStyle: 'italic' };
const BOLD_GLYPH: CSSProperties = { fontWeight: 700 };
const UNDERLINE_GLYPH: CSSProperties = { textDecoration: 'underline' };

function glyphStyle(id: string): CSSProperties | undefined {
  if (id === 'bold') return BOLD_GLYPH;
  if (id === 'italic') return ITALIC_GLYPH;
  if (id === 'underline') return UNDERLINE_GLYPH;
  return undefined;
}

export interface ToolbarProps {
  /** Reaches the live API (set after `onReady`); read lazily on click so the
   *  toolbar can render before the editor finishes booting. */
  getApi: () => CasualSheetsAPI | null;
}

export function Toolbar({ getApi }: ToolbarProps) {
  return (
    <div style={BAR_STYLE} data-testid="casual-sheets-toolbar" role="toolbar" aria-label="Editor">
      {ACTIONS.map((a) => (
        <button
          key={a.id}
          type="button"
          title={a.label}
          aria-label={a.label}
          data-action={a.id}
          style={BTN_STYLE}
          // Mouse-down (not click) so the grid's active selection isn't lost to
          // a focus change before the command runs.
          onMouseDown={(e) => {
            e.preventDefault();
            void getApi()?.executeCommand(a.command);
          }}
        >
          <span style={glyphStyle(a.id)}>{a.glyph}</span>
        </button>
      ))}
    </div>
  );
}
