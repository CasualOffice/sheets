import { useEffect, useState } from 'react';
import { useUniverAPI } from '../use-univer';
import { useUI } from '../use-ui';
import { useOutline } from '../outline/outline-context';
import type { OutlineGroup } from '../outline/types';
import { Icon } from './Icon';
import { Tooltip } from './Tooltip';

/**
 * Side panel that lists outline groups on the active sheet and exposes one-
 * click collapse / expand / remove per group. Toggled from Data → Outline.
 *
 * We don't try to render Excel-style +/− buttons in the grid gutter — the
 * gutter is Univer's canvas territory and overlaying DOM there is brittle.
 * The panel is a strict superset of that functionality (you can see every
 * group at once, with its range and state) and avoids fighting the canvas.
 */
export function OutlinePanel() {
  const api = useUniverAPI();
  const ui = useUI();
  const outline = useOutline();
  const [activeSheetId, setActiveSheetId] = useState<string | null>(null);

  // Track which sheet is active so the panel always lists the user's
  // current focus. Tied to Univer's command bus rather than polling.
  useEffect(() => {
    if (!api) return;
    const read = () => {
      const id = api.getActiveWorkbook()?.getActiveSheet()?.getSheetId() ?? null;
      setActiveSheetId(id);
    };
    read();
    const disp = api.addEvent(api.Event.CommandExecuted, (e) => {
      const id = (e as { id?: string }).id;
      if (
        id === 'sheet.operation.set-worksheet-activate' ||
        id === 'doc.command-replace-snapshot'
      ) {
        read();
      }
    });
    return () => disp.dispose();
  }, [api]);

  const sheet = activeSheetId ? outline.getSheet(activeSheetId) : { rows: [], cols: [] };
  const empty = sheet.rows.length === 0 && sheet.cols.length === 0;

  return (
    <aside className="side-panel outline-panel" data-testid="outline-panel">
      <header className="side-panel__header">
        <Icon name="account_tree" size="sm" />
        <h2 className="side-panel__title">Outline</h2>
        <Tooltip label="Close outline panel">
          <button
            type="button"
            className="side-panel__close"
            data-testid="outline-panel-close"
            aria-label="Close outline panel"
            onClick={ui.toggleOutlinePanel}
          >
            <Icon name="close" size="sm" />
          </button>
        </Tooltip>
      </header>
      <div className="outline-panel__body">
        {empty ? (
          <div className="outline-panel__empty">
            No groups on this sheet. Select rows or columns and use
            <strong> Data → Group rows / Group columns</strong> to create one.
          </div>
        ) : (
          <>
            {sheet.rows.length > 0 && (
              <OutlineSection
                title="Row groups"
                axis="rows"
                groups={sheet.rows}
                sheetId={activeSheetId!}
              />
            )}
            {sheet.cols.length > 0 && (
              <OutlineSection
                title="Column groups"
                axis="cols"
                groups={sheet.cols}
                sheetId={activeSheetId!}
              />
            )}
          </>
        )}
      </div>
    </aside>
  );
}

function OutlineSection({
  title,
  axis,
  groups,
  sheetId,
}: {
  title: string;
  axis: 'rows' | 'cols';
  groups: OutlineGroup[];
  sheetId: string;
}) {
  const outline = useOutline();
  const sorted = [...groups].sort((a, b) => a.start - b.start);
  return (
    <section className="outline-panel__section" data-testid={`outline-section-${axis}`}>
      <h4 className="outline-panel__section-title">{title}</h4>
      <ul className="outline-panel__list">
        {sorted.map((g) => (
          <li className="outline-panel__row" key={g.id} data-testid={`outline-row-${axis}-${g.id}`}>
            <Tooltip label={g.collapsed ? 'Expand' : 'Collapse'}>
              <button
                type="button"
                className="outline-panel__toggle"
                data-testid={`outline-toggle-${axis}-${g.id}`}
                aria-label={g.collapsed ? 'Expand' : 'Collapse'}
                aria-pressed={g.collapsed}
                onClick={() => outline.setCollapsed(sheetId, axis, g.id, !g.collapsed)}
              >
                <Icon name={g.collapsed ? 'chevron_right' : 'expand_more'} size="sm" />
              </button>
            </Tooltip>
            <span className="outline-panel__range">{formatRange(axis, g)}</span>
            <Tooltip label="Remove group">
              <button
                type="button"
                className="outline-panel__remove"
                data-testid={`outline-remove-${axis}-${g.id}`}
                aria-label="Remove group"
                onClick={() => outline.removeGroup(sheetId, axis, g.id)}
              >
                <Icon name="close" size="sm" />
              </button>
            </Tooltip>
          </li>
        ))}
      </ul>
    </section>
  );
}

function formatRange(axis: 'rows' | 'cols', g: OutlineGroup): string {
  if (axis === 'rows') {
    return g.start === g.end ? `Row ${g.start + 1}` : `Rows ${g.start + 1}–${g.end + 1}`;
  }
  const letters = (n: number) => {
    let m = n + 1;
    let out = '';
    while (m > 0) {
      const r = (m - 1) % 26;
      out = String.fromCharCode(65 + r) + out;
      m = Math.floor((m - 1) / 26);
    }
    return out;
  };
  return g.start === g.end
    ? `Column ${letters(g.start)}`
    : `Columns ${letters(g.start)}–${letters(g.end)}`;
}
