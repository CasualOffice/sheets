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

/**
 * Dialog — the SDK chrome's modal primitive for `<CasualSheets chrome>`.
 *
 * Ported from `apps/web/src/shell/Dialog.tsx`, but decoupled from the app: it
 * uses ONLY React + the chrome's `--cs-chrome-*` token approach (inline styles
 * with DS-token-backed CSS-var fallbacks, the same pattern as FindReplace), so
 * it renders standalone and themes light/dark via the `data-theme` wrapper
 * `<CasualSheets>` already sets. No app CSS classes, no app context.
 *
 * - Click-backdrop and Escape close it.
 * - Body scroll is locked while open.
 * - Focus is trapped (Tab/Shift+Tab cycle within the dialog) and restored to
 *   the opener on close — the WAI-ARIA dialog pattern.
 * - Portaled to `document.body` so it escapes the chrome's stacking contexts
 *   and the Univer canvas can't intercept clicks meant for the dialog.
 */

import { useEffect, useRef, type CSSProperties, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { Icon } from './Icon';

export interface DialogProps {
  /** Heading shown in the dialog header. */
  title: string;
  /** Close request (backdrop click, Escape, the × button, or a footer action). */
  onClose: () => void;
  /** Optional footer node — typically Cancel / primary-action buttons. */
  footer?: ReactNode;
  children: ReactNode;
  /** Max width of the dialog card. Default 460. */
  width?: number;
  ['data-testid']?: string;
}

const FOCUSABLE_SELECTOR =
  'a[href], area[href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), button:not([disabled]), iframe, object, embed, [tabindex]:not([tabindex="-1"]), [contenteditable="true"]';

const BACKDROP_STYLE: CSSProperties = {
  position: 'fixed',
  inset: 0,
  zIndex: 1200,
  display: 'flex',
  alignItems: 'flex-start',
  justifyContent: 'center',
  padding: '8vh 16px 16px',
  background: 'rgba(15, 23, 42, 0.35)',
  overflow: 'auto',
};

const CARD_STYLE: CSSProperties = {
  width: '100%',
  background: 'var(--cs-chrome-input-bg, #ffffff)',
  color: 'var(--cs-chrome-fg, #201f1e)',
  borderRadius: 10,
  border: '1px solid var(--cs-chrome-border, #e6e9ee)',
  boxShadow: '0 18px 48px rgba(0, 0, 0, 0.28)',
  font: 'inherit',
  fontSize: 13,
  display: 'flex',
  flexDirection: 'column',
  maxHeight: '84vh',
};

const HEADER_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  gap: 8,
  padding: '12px 16px',
  borderBottom: '1px solid var(--cs-chrome-border, #edeff3)',
};

const TITLE_STYLE: CSSProperties = {
  margin: 0,
  fontSize: 15,
  fontWeight: 600,
  color: 'var(--cs-chrome-fg, #201f1e)',
};

const CLOSE_STYLE: CSSProperties = {
  width: 28,
  height: 28,
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  border: 'none',
  borderRadius: 6,
  background: 'transparent',
  color: 'var(--cs-chrome-muted, #605e5c)',
  cursor: 'pointer',
  padding: 0,
};

const BODY_STYLE: CSSProperties = {
  padding: 16,
  overflow: 'auto',
};

const FOOTER_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'flex-end',
  gap: 8,
  padding: '12px 16px',
  borderTop: '1px solid var(--cs-chrome-border, #edeff3)',
};

export function Dialog({ title, onClose, footer, children, width = 460, ...rest }: DialogProps) {
  const dialogRef = useRef<HTMLDivElement>(null);
  const openerRef = useRef<HTMLElement | null>(null);

  useEffect(() => {
    // Remember who had focus before us so we can restore it on close.
    openerRef.current = (document.activeElement as HTMLElement) ?? null;

    // Lock body scroll; pad for the disappearing scrollbar so the grid doesn't
    // shift when the dialog opens.
    const prevOverflow = document.body.style.overflow;
    const prevPaddingRight = document.body.style.paddingRight;
    const scrollbarW = window.innerWidth - document.documentElement.clientWidth;
    document.body.style.overflow = 'hidden';
    if (scrollbarW > 0) document.body.style.paddingRight = `${scrollbarW}px`;

    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation();
        onClose();
        return;
      }
      if (e.key !== 'Tab') return;
      const root = dialogRef.current;
      if (!root) return;
      const focusables = Array.from(root.querySelectorAll<HTMLElement>(FOCUSABLE_SELECTOR)).filter(
        (el) => !el.hasAttribute('disabled') && el.offsetParent !== null,
      );
      if (focusables.length === 0) return;
      const first = focusables[0];
      const last = focusables[focusables.length - 1];
      const active = document.activeElement as HTMLElement | null;
      if (e.shiftKey && (active === first || !root.contains(active))) {
        e.preventDefault();
        last.focus();
      } else if (!e.shiftKey && (active === last || !root.contains(active))) {
        e.preventDefault();
        first.focus();
      }
    };
    document.addEventListener('keydown', onKey);

    const first = dialogRef.current?.querySelector<HTMLElement>(FOCUSABLE_SELECTOR);
    first?.focus();

    return () => {
      document.removeEventListener('keydown', onKey);
      document.body.style.overflow = prevOverflow;
      document.body.style.paddingRight = prevPaddingRight;
      const opener = openerRef.current;
      if (opener && document.contains(opener)) opener.focus();
    };
  }, [onClose]);

  return createPortal(
    <div
      style={BACKDROP_STYLE}
      data-testid="cs-dialog-backdrop"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        style={{ ...CARD_STYLE, maxWidth: width }}
        role="dialog"
        aria-modal="true"
        aria-label={title}
        ref={dialogRef}
        data-testid={rest['data-testid']}
      >
        <div style={HEADER_STYLE}>
          <h2 style={TITLE_STYLE}>{title}</h2>
          <button
            type="button"
            style={CLOSE_STYLE}
            data-testid="cs-dialog-close"
            aria-label="Close"
            onClick={onClose}
          >
            <Icon name="close" size={18} />
          </button>
        </div>
        <div style={BODY_STYLE}>{children}</div>
        {footer && <div style={FOOTER_STYLE}>{footer}</div>}
      </div>
    </div>,
    document.body,
  );
}
