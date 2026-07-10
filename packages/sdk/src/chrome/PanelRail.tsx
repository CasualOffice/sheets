/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License").
 */

/**
 * Right-edge vertical rail of panel-toggle buttons (mirrors the standalone
 * app's PanelRail, but registry-driven and self-contained in the SDK). Built-in
 * panels come first, then any host `extensions.panels`. Clicking a button
 * toggles that panel through the shared panel store; the open panel's body is
 * rendered to the left by PanelHost.
 */
import type { CSSProperties } from 'react';

import type { ChromeExtensions } from './extensions';
import { Icon } from './Icon';
import { usePanels } from './panel-context';
import { BUILT_IN_PANELS } from './panel-registry';

const railStyle: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  gap: 2,
  padding: '6px 4px',
  flex: '0 0 auto',
  borderLeft: '1px solid var(--cs-chrome-border, #edeff3)',
  background: 'var(--cs-chrome-bg, #eef1f5)',
};

export interface RailEntry {
  id: string;
  label: string;
  icon: string;
}

/** Merge built-in panels with host-supplied ones into a single rail list. */
export function railEntries(extensions?: ChromeExtensions): RailEntry[] {
  const built = BUILT_IN_PANELS.map((p) => ({ id: p.id, label: p.label, icon: p.icon }));
  const host = (extensions?.panels ?? []).map((p) => ({
    id: p.id,
    label: p.title,
    icon: p.railIcon,
  }));
  return [...built, ...host];
}

export function PanelRail({ extensions }: { extensions?: ChromeExtensions }) {
  const panels = usePanels();
  const entries = railEntries(extensions);
  if (entries.length === 0) return null;

  return (
    <aside style={railStyle} data-testid="cs-panel-rail" aria-label="Panels" role="toolbar">
      {entries.map((e) => {
        const pressed = panels.openPanelId === e.id;
        return (
          <button
            key={e.id}
            type="button"
            data-testid={`cs-panel-rail-${e.id}`}
            aria-pressed={pressed}
            aria-label={pressed ? `Hide ${e.label}` : e.label}
            title={pressed ? `Hide ${e.label}` : e.label}
            onClick={() => panels.toggle(e.id)}
            style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              width: 30,
              height: 30,
              border: 'none',
              borderRadius: 6,
              cursor: 'pointer',
              color: pressed
                ? 'var(--cs-chrome-active-fg, #0e7490)'
                : 'var(--cs-chrome-fg, #201f1e)',
              background: pressed
                ? 'var(--cs-chrome-active, rgba(14,116,144,0.11))'
                : 'transparent',
            }}
          >
            <Icon name={e.icon} size={18} />
          </button>
        );
      })}
    </aside>
  );
}
