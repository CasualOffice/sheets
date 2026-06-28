/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

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
      { label: 'Find', combo: 'Ctrl+F' },
      { label: 'Find & replace', combo: 'Ctrl+H' },
      { label: 'Go to / Name Box', combo: 'Ctrl+G' },
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
      { label: 'Fill down', combo: 'Ctrl+D' },
      { label: 'Fill right', combo: 'Ctrl+R' },
      { label: 'Copy from cell above', combo: "Ctrl+'" },
      { label: 'Flash Fill', combo: 'Ctrl+E' },
      { label: 'Insert cells / rows', combo: 'Ctrl++' },
      { label: 'Delete cells / rows', combo: 'Ctrl+-' },
      { label: 'Insert link', combo: 'Ctrl+K' },
      { label: 'Insert comment', combo: 'Shift+F2' },
    ],
  },
  {
    heading: 'Navigation & selection',
    rows: [
      { label: 'Go to start of sheet', combo: 'Ctrl+Home' },
      { label: 'Go to end of data', combo: 'Ctrl+End' },
      { label: 'Previous sheet tab', combo: 'Ctrl+PageUp' },
      { label: 'Next sheet tab', combo: 'Ctrl+PageDown' },
      { label: 'Select entire column', combo: 'Ctrl+Space' },
      { label: 'Select entire row', combo: 'Shift+Space' },
      { label: 'Add to selection mode', combo: 'Shift+F8' },
      { label: 'Context menu for cell', combo: 'Shift+F10' },
    ],
  },
  {
    heading: 'Formatting',
    rows: [
      { label: 'Format cells…', combo: 'Ctrl+1' },
      { label: 'Toggle formula view', combo: 'Ctrl+`' },
      { label: 'Insert date', combo: 'Ctrl+;' },
      { label: 'Insert time', combo: 'Ctrl+Shift+:' },
      { label: 'Format as table', combo: 'Ctrl+L' },
      { label: 'Outside border', combo: 'Ctrl+Shift+7' },
      { label: 'Grow font size', combo: 'Ctrl+Shift+>' },
      { label: 'Shrink font size', combo: 'Ctrl+Shift+<' },
      { label: 'Hide row', combo: 'Ctrl+9' },
      { label: 'Unhide row', combo: 'Ctrl+Shift+9' },
      { label: 'Hide column', combo: 'Ctrl+0' },
      { label: 'Unhide column', combo: 'Ctrl+Shift+0' },
    ],
  },
  {
    heading: 'Formulas & data',
    rows: [
      { label: 'AutoSum', combo: 'Alt+=' },
      { label: 'Insert function…', combo: 'Shift+F3' },
      { label: 'Recalculate', combo: 'F9' },
      { label: 'Toggle filter', combo: 'Ctrl+Shift+L' },
      { label: 'Trace precedents', combo: 'Ctrl+[' },
      { label: 'Trace dependents', combo: 'Ctrl+]' },
    ],
  },
  {
    heading: 'Insert & sheets',
    rows: [
      { label: 'Insert sheet', combo: 'Shift+F11' },
      { label: 'Insert chart', combo: 'Alt+F1' },
      { label: 'Name Manager', combo: 'Ctrl+F3' },
      { label: 'Show pivot details', combo: 'Ctrl+Shift+D' },
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
                color: 'var(--color-text-secondary, #475569)',
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
                    color: 'var(--color-text, #0f172a)',
                    gap: 12,
                  }}
                >
                  <span style={{ flex: 1, minWidth: 0 }}>{row.label}</span>
                  <kbd
                    style={{
                      fontFamily:
                        'ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace',
                      fontSize: 12,
                      background: 'var(--color-surface-alt, #f1f5f9)',
                      border: '1px solid var(--color-border, #e2e8f0)',
                      borderRadius: 4,
                      padding: '2px 7px',
                      color: 'var(--color-text, #0f172a)',
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
