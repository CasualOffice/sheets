import { useMemo } from 'react';
import { Dialog } from './Dialog';
import { formatShortcut } from './shortcut-format';

/**
 * Keyboard-shortcut cheat sheet — UX_AUDIT.md §2.16 / Phase 3 #12.
 *
 * Opens via `?` (matches Google Docs / Gmail / Slack) or `Ctrl+/`
 * (matches Excel for the Web). Lists the most-used keyboard
 * affordances grouped the way users mentally chunk them: Editing,
 * Navigation, Formatting, View, App. Not exhaustive — the MenuBar
 * itself stays the canonical surface for the long tail.
 *
 * Cross-platform: every shortcut string is the canonical
 * `Ctrl+<key>` form, run through `formatShortcut(navigator.platform)`
 * so Mac users see `⌘` and Win/Linux users see `Ctrl`. Same util
 * MenuBar uses, so the cheat sheet and the menu agree.
 */
type Props = {
  onClose: () => void;
};

interface ShortcutRow {
  label: string;
  /** Canonical form, e.g. `Ctrl+S`. Mac/Win rendering handled in `formatShortcut`. */
  combo: string;
}

interface ShortcutGroup {
  heading: string;
  rows: ShortcutRow[];
}

const GROUPS: ShortcutGroup[] = [
  {
    heading: 'Essentials',
    rows: [
      { label: 'New workbook', combo: 'Ctrl+N' },
      { label: 'Open…', combo: 'Ctrl+O' },
      { label: 'Save', combo: 'Ctrl+S' },
      { label: 'Print', combo: 'Ctrl+P' },
      { label: 'Find & replace', combo: 'Ctrl+F' },
      { label: 'Quick actions / Tell Me', combo: 'Alt+Q' },
      { label: 'Command palette', combo: 'Ctrl+Shift+P' },
      { label: 'Keyboard shortcuts (this dialog)', combo: 'Ctrl+/' },
    ],
  },
  {
    heading: 'Editing',
    rows: [
      { label: 'Undo', combo: 'Ctrl+Z' },
      { label: 'Redo', combo: 'Ctrl+Y' },
      { label: 'Cut', combo: 'Ctrl+X' },
      { label: 'Copy', combo: 'Ctrl+C' },
      { label: 'Paste', combo: 'Ctrl+V' },
      { label: 'Paste values only', combo: 'Ctrl+Shift+V' },
      { label: 'Paste special…', combo: 'Ctrl+Alt+V' },
      { label: 'Edit cell in place', combo: 'F2' },
      { label: 'Flash Fill', combo: 'Ctrl+E' },
      { label: 'Insert link', combo: 'Ctrl+K' },
      { label: 'Insert comment', combo: 'Shift+F2' },
    ],
  },
  {
    heading: 'Navigation',
    rows: [
      { label: 'Go to start of sheet', combo: 'Ctrl+Home' },
      { label: 'Go to end of data', combo: 'Ctrl+End' },
      { label: 'Previous sheet tab', combo: 'Ctrl+PageUp' },
      { label: 'Next sheet tab', combo: 'Ctrl+PageDown' },
      { label: 'Select entire column', combo: 'Ctrl+Space' },
      { label: 'Select entire row', combo: 'Shift+Space' },
    ],
  },
  {
    heading: 'Formatting',
    rows: [
      { label: 'Format cells…', combo: 'Ctrl+1' },
      { label: "Toggle formula view", combo: 'Ctrl+`' },
      { label: 'Insert date', combo: 'Ctrl+;' },
      { label: 'Insert time', combo: 'Ctrl+Shift+:' },
      { label: 'Format as table', combo: 'Ctrl+L' },
      { label: 'Hide row', combo: 'Ctrl+9' },
      { label: 'Unhide row', combo: 'Ctrl+Shift+9' },
      { label: 'Hide column', combo: 'Ctrl+0' },
      { label: 'Unhide column', combo: 'Ctrl+Shift+0' },
    ],
  },
  {
    heading: 'View',
    rows: [
      { label: 'Zoom in', combo: 'Ctrl++' },
      { label: 'Zoom out', combo: 'Ctrl+-' },
      { label: 'Toggle full screen', combo: 'Shift+F11' },
      { label: 'Define name', combo: 'Ctrl+F3' },
      { label: 'Refresh data', combo: 'Ctrl+Shift+D' },
    ],
  },
];

export function KeyboardShortcutsDialog({ onClose }: Props) {
  // Detect platform inside the component so the rendering matches what
  // the user actually sees (formatShortcut already handles this for
  // MenuBar; we reuse it here for consistency).
  const platform = useMemo(() => navigator.platform, []);
  return (
    <Dialog
      title="Keyboard shortcuts"
      onClose={onClose}
      data-testid="keyboard-shortcuts-dialog"
      footer={
        <button
          type="button"
          className="btn-primary"
          data-testid="keyboard-shortcuts-close"
          onClick={onClose}
        >
          Close
        </button>
      }
    >
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))',
          gap: '24px 32px',
          minWidth: 600,
          maxWidth: 760,
        }}
      >
        {GROUPS.map((group) => (
          <section key={group.heading}>
            <h3
              style={{
                margin: '0 0 8px',
                fontSize: 13,
                textTransform: 'uppercase',
                letterSpacing: '0.04em',
                color: '#475569',
                fontWeight: 600,
              }}
            >
              {group.heading}
            </h3>
            <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
              {group.rows.map((row) => (
                <li
                  key={row.combo}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    padding: '5px 0',
                    fontSize: 13,
                    color: '#0f172a',
                    gap: 12,
                  }}
                >
                  <span style={{ flex: 1, minWidth: 0 }}>{row.label}</span>
                  <kbd
                    style={{
                      fontFamily:
                        'ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace',
                      fontSize: 12,
                      background: '#f1f5f9',
                      border: '1px solid #e2e8f0',
                      borderRadius: 4,
                      padding: '2px 7px',
                      color: '#0f172a',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {formatShortcut(row.combo, platform)}
                  </kbd>
                </li>
              ))}
            </ul>
          </section>
        ))}
      </div>
    </Dialog>
  );
}
