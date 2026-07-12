/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License").
 */

/**
 * Shared chrome for the side panels — a consistent, polished header (icon +
 * title + optional count/actions + hover-capable close) and small building
 * blocks, so every panel (Tables / Charts / Pivot / Comments / History) looks
 * the same. Styling maps to the design-system tokens the host loads
 * (`--color-*`), falling back to the chrome vars, then a hardcoded default —
 * so it looks right standalone and themed inside a host.
 */
import { useState, type CSSProperties, type ReactNode } from 'react';

import { Icon } from './Icon';

const MUTED = 'var(--color-text-secondary, var(--cs-chrome-muted, #605e5c))';
const DIVIDER = 'var(--color-divider, var(--cs-chrome-border, #edeff3))';

const headerStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 10,
  padding: '10px 14px',
  flex: '0 0 auto',
  background: 'var(--color-surface, var(--cs-chrome-input-bg, #ffffff))',
  borderBottom: `1px solid ${DIVIDER}`,
};

const titleStyle: CSSProperties = {
  fontWeight: 600,
  fontSize: 14,
  flex: 1,
  minWidth: 0,
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap',
};

const countStyle: CSSProperties = {
  fontSize: 11,
  fontWeight: 600,
  minWidth: 18,
  height: 18,
  padding: '0 5px',
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  borderRadius: 9,
  background: 'var(--color-hover, rgba(15,23,42,0.06))',
  color: MUTED,
};

/** A ghost icon button with a real hover state (inline styles can't do :hover). */
export function IconButton({
  name,
  label,
  onClick,
  size = 18,
}: {
  name: string;
  label: string;
  onClick: () => void;
  size?: number;
}) {
  const [hover, setHover] = useState(false);
  return (
    <button
      type="button"
      aria-label={label}
      title={label}
      onClick={onClick}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      style={{
        flex: '0 0 auto',
        display: 'inline-flex',
        alignItems: 'center',
        justifyContent: 'center',
        width: 26,
        height: 26,
        border: 'none',
        borderRadius: 6,
        cursor: 'pointer',
        color: hover ? 'var(--color-text, var(--cs-chrome-fg, #201f1e))' : MUTED,
        background: hover ? 'var(--color-hover, rgba(15,23,42,0.06))' : 'transparent',
      }}
    >
      <Icon name={name} size={size} />
    </button>
  );
}

export function PanelHeader({
  icon,
  title,
  count,
  onClose,
  actions,
}: {
  icon: string;
  title: string;
  count?: number;
  onClose: () => void;
  /** Optional trailing actions (e.g. an add button) shown before the close. */
  actions?: ReactNode;
}) {
  return (
    <header style={headerStyle}>
      <Icon name={icon} size={18} style={{ color: MUTED, flex: '0 0 auto' }} />
      <span style={titleStyle}>{title}</span>
      {count != null && count > 0 && <span style={countStyle}>{count}</span>}
      {actions}
      <IconButton name="close" label={`Close ${title} panel`} onClick={onClose} />
    </header>
  );
}

/** Centered empty-state block for panels with no content yet. */
export function PanelEmpty({
  icon,
  title,
  children,
  testId,
}: {
  icon: string;
  title: string;
  children?: ReactNode;
  testId?: string;
}) {
  return (
    <div
      data-testid={testId}
      style={{
        textAlign: 'center',
        padding: '32px 20px',
        color: MUTED,
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        gap: 6,
      }}
    >
      <Icon name={icon} size={40} style={{ opacity: 0.35 }} />
      <div style={{ fontWeight: 600, fontSize: 14, color: 'var(--color-text, inherit)' }}>{title}</div>
      {children && <div style={{ fontSize: 13, lineHeight: 1.5 }}>{children}</div>}
    </div>
  );
}
