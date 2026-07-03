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

import { useUI } from '../use-ui';
import { Icon } from './Icon';
import { isDesktop } from '../desk-bridge-bootstrap';
import { hasCollabServer } from '../ai/transport';

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
  // AI panel only available on desktop or when a collab server is configured.
  // In plain web without either, the DirectTransport API-key flow is hidden.
  const showAiButton = isDesktop() || hasCollabServer();

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
        id="pivot"
        label={ui.pivotPanelVisible ? 'Hide PivotTable Fields' : 'PivotTable Fields'}
        icon="pivot_table_chart"
        pressed={ui.pivotPanelVisible}
        onClick={ui.togglePivotPanel}
      />
      <RailButton
        id="watch"
        label={ui.watchPanelVisible ? 'Hide Watch Window' : 'Watch Window'}
        icon="visibility"
        pressed={ui.watchPanelVisible}
        onClick={ui.toggleWatchPanel}
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
      {showAiButton && (
        <RailButton
          id="ai"
          label={ui.aiPanelVisible ? 'Hide AI' : 'AI'}
          icon="auto_awesome"
          pressed={ui.aiPanelVisible}
          onClick={ui.toggleAiPanel}
        />
      )}
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
