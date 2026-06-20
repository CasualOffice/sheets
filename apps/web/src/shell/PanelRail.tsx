import { useUI } from '../use-ui';
import { Icon } from './Icon';

/**
 * Right-edge vertical rail of panel-toggle buttons. Replaces the
 * panel-toggle group that used to live in the toolbar — keeping panel
 * toggles in one always-visible spot makes them easier to find and
 * removes the duplication we had between the toolbar's Panels group
 * and the View menu items.
 *
 * Each button shows its panel's pressed state. Clicking opens the
 * panel (its body renders to the left of this rail, in the same
 * grid-row flex line). The rail itself never collapses — even with
 * every panel closed, the icons remain accessible.
 *
 * Every entry — Comments included — is now a React panel sharing the
 * `.side-panel` shell + the mutex, so each shows a real pressed state.
 */
export function PanelRail() {
  const ui = useUI();

  return (
    <aside className="panel-rail" data-testid="panel-rail" aria-label="Panels">
      <RailButton
        id="tables"
        label={ui.tablesPanelVisible ? 'Hide Tables' : 'Tables'}
        icon="table"
        pressed={ui.tablesPanelVisible}
        onClick={ui.toggleTablesPanel}
      />
      <RailButton
        id="charts"
        label={ui.chartsPanelVisible ? 'Hide Charts' : 'Charts'}
        icon="analytics"
        pressed={ui.chartsPanelVisible}
        onClick={ui.toggleChartsPanel}
      />
      <RailButton
        id="outline"
        label={ui.outlinePanelVisible ? 'Hide Outline' : 'Outline'}
        icon="format_list_bulleted"
        pressed={ui.outlinePanelVisible}
        onClick={ui.toggleOutlinePanel}
      />
      <RailButton
        id="comments"
        label={ui.commentsPanelVisible ? 'Hide Comments' : 'Comments'}
        icon="forum"
        pressed={ui.commentsPanelVisible}
        onClick={ui.toggleCommentsPanel}
      />
      <RailButton
        id="history"
        label={ui.historyPanelVisible ? 'Hide History' : 'History'}
        icon="history"
        pressed={ui.historyPanelVisible}
        onClick={ui.toggleHistoryPanel}
      />
    </aside>
  );
}

function RailButton({
  id,
  label,
  icon,
  pressed,
  onClick,
}: {
  id: string;
  label: string;
  icon: string;
  pressed?: boolean;
  onClick?: () => void;
}) {
  return (
    <button
      type="button"
      className={`panel-rail__btn${pressed ? ' panel-rail__btn--active' : ''}`}
      data-testid={`panel-rail-${id}`}
      aria-pressed={pressed ?? undefined}
      aria-label={label}
      title={label}
      onClick={onClick}
    >
      <Icon name={icon} />
    </button>
  );
}
