/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License").
 */

/**
 * Built-in side panels the SDK chrome ships (mirrors `BUILT_IN_DIALOGS`). Each
 * entry gets a rail button + a body rendered by the panel host. Hosts can add
 * more via `extensions.panels` (see extensions.ts) — those are merged after
 * these on the rail.
 */
import { lazy, type ComponentType } from 'react';

import type { PanelComponentProps } from './extensions';
import { TablesPanel } from './TablesPanel';
import { PivotFieldsPanel } from './PivotFieldsPanel';
import { CommentsPanel } from './CommentsPanel';
import { HistoryPanel } from './HistoryPanel';
// Charts panel pulls echarts — lazy so it (and echarts) only load when the
// panel is first opened. PanelHost renders panels inside a Suspense boundary.
const ChartsPanel = lazy(() =>
  import('../charts/ChartsPanel').then((m) => ({ default: m.ChartsPanel })),
);

export interface BuiltInPanel {
  /** Stable id (rail `data-testid` `cs-panel-rail-<id>`, mutex key). */
  id: string;
  /** Rail button tooltip / aria-label. */
  label: string;
  /** Material Symbols icon name for the rail button. */
  icon: string;
  /** The panel body; receives `{ api, onClose }`. */
  component: ComponentType<PanelComponentProps>;
}

export const BUILT_IN_PANELS: BuiltInPanel[] = [
  { id: 'tables', label: 'Tables', icon: 'table', component: TablesPanel },
  {
    id: 'pivot',
    label: 'PivotTable Fields',
    icon: 'pivot_table_chart',
    component: PivotFieldsPanel,
  },
  { id: 'charts', label: 'Charts', icon: 'analytics', component: ChartsPanel },
  { id: 'comments', label: 'Comments', icon: 'forum', component: CommentsPanel },
  { id: 'history', label: 'History', icon: 'history', component: HistoryPanel },
];
