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
import type { ComponentType } from 'react';

import type { PanelComponentProps } from './extensions';
import { TablesPanel } from './TablesPanel';

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
];
