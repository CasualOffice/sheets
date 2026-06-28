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

import { useEffect, useRef, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { Icon } from './Icon';

type Props = {
  title: string;
  onClose: () => void;
  footer?: ReactNode;
  children: ReactNode;
  ['data-testid']?: string;
};

const FOCUSABLE_SELECTOR =
  'a[href], area[href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), button:not([disabled]), iframe, object, embed, [tabindex]:not([tabindex="-1"]), [contenteditable="true"]';

/**
 * Generic modal dialog with proper focus trap + body-scroll lock.
 *
 * - Click-backdrop and Escape close it.
 * - Body scroll is locked while open (grid underneath used to scroll
 *   when scrolling inside a dialog hit the end of its body).
 * - Focus is trapped: Tab cycles through focusable descendants instead
 *   of escaping to the page below.
 * - Focus is returned to the element that opened the dialog when the
 *   dialog closes, so keyboard users don't get dumped at `<body>`.
 */
export function Dialog({ title, onClose, footer, children, ...rest }: Props) {
  const dialogRef = useRef<HTMLDivElement>(null);
  const openerRef = useRef<HTMLElement | null>(null);

  useEffect(() => {
    // Remember who had focus before us so we can restore it on close.
    // Most callers are toolbar/menu buttons — falling back to body
    // leaves keyboard users disoriented after Esc.
    openerRef.current = (document.activeElement as HTMLElement) ?? null;

    // Lock body scroll. Padding-right compensates for the disappearing
    // scrollbar so the grid doesn't shift when the dialog opens.
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
      const focusables = Array.from(
        root.querySelectorAll<HTMLElement>(FOCUSABLE_SELECTOR),
      ).filter((el) => !el.hasAttribute('disabled') && el.offsetParent !== null);
      if (focusables.length === 0) return;
      const first = focusables[0];
      const last = focusables[focusables.length - 1];
      const active = document.activeElement as HTMLElement | null;
      // Shift+Tab on the first element wraps to last; Tab on the last
      // wraps to first. Matches the WAI-ARIA dialog pattern.
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
      // Restore focus, but only if it's still in the document — the
      // opener might have been unmounted (e.g. a context menu that
      // triggered the dialog has already gone away).
      const opener = openerRef.current;
      if (opener && document.contains(opener)) opener.focus();
    };
  }, [onClose]);

  // Always portal modals to document.body. Without this, dialogs
  // opened from inside a sibling that uses `overflow:hidden` (like
  // ChartLayer for the chart context menu) get clipped to that
  // parent's bounds and any chart-canvas underneath intercepts the
  // clicks meant for the dialog buttons.
  return createPortal(
    <div
      className="dialog-backdrop"
      data-testid="dialog-backdrop"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        className="dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="dialog-title"
        ref={dialogRef}
        data-testid={rest['data-testid']}
      >
        <div className="dialog__header">
          <h2 className="dialog__title" id="dialog-title">
            {title}
          </h2>
          <button
            type="button"
            className="btn btn--icon"
            data-testid="dialog-close"
            aria-label="Close"
            onClick={onClose}
          >
            <Icon name="close" size="sm" />
          </button>
        </div>
        <div className="dialog__body">{children}</div>
        {footer && <div className="dialog__footer">{footer}</div>}
      </div>
    </div>,
    document.body,
  );
}
