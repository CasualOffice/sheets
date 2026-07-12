/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License").
 */

/**
 * Renders the currently-open side panel's body (built-in or host-supplied),
 * to the left of the rail. Returns null when no panel is open so the grid keeps
 * the full width. Each panel body receives `{ api, onClose }`.
 */
import { Suspense, type CSSProperties } from 'react';

import type { CasualSheetsAPI } from '../sheets/api';
import type { ChromeExtensions } from './extensions';
import { usePanels } from './panel-context';
import { BUILT_IN_PANELS } from './panel-registry';

const asideStyle: CSSProperties = {
  flex: '0 0 auto',
  width: 320,
  minWidth: 0,
  height: '100%',
  overflow: 'auto',
  borderLeft: '1px solid var(--color-divider, var(--cs-chrome-border, #edeff3))',
  background: 'var(--color-surface-alt, var(--cs-chrome-input-bg, #ffffff))',
  color: 'var(--color-text, var(--cs-chrome-fg, #201f1e))',
  fontSize: 13,
};

export function PanelHost({
  api,
  extensions,
}: {
  api: CasualSheetsAPI | null;
  extensions?: ChromeExtensions;
}) {
  const panels = usePanels();
  const openId = panels.openPanelId;
  if (!openId || !api) return null;

  const builtIn = BUILT_IN_PANELS.find((p) => p.id === openId);
  const host = extensions?.panels?.find((p) => p.id === openId);
  const Body = builtIn?.component ?? host?.component;
  if (!Body) return null;

  return (
    <aside style={asideStyle} data-testid={`cs-panel-${openId}`}>
      {/* Panels may be lazy (e.g. Charts pulls echarts) — Suspense keeps the
          rest of the chrome painted while the chunk loads. */}
      <Suspense fallback={null}>
        <Body api={api} onClose={panels.close} />
      </Suspense>
    </aside>
  );
}
