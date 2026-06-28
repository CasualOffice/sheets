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

import { useToast, useToastList } from './toast-context';

/**
 * Renders the active toast stack. Designed to sit at the app root,
 * AFTER all other shell elements — fixed-positioned bottom-right so
 * it doesn't compete with the formula bar or the mobile action bar
 * for vertical space.
 *
 * No portal — the container is plain absolute inside <body>'s flex
 * tree. Toast stacking is `display: flex; flex-direction: column-
 * reverse; gap;` so the newest toast lands at the bottom (closest
 * to the user's eye) without re-flowing older toasts up the screen.
 *
 * Per-toast a11y:
 *   - role="status" + aria-live="polite" for info/success
 *   - role="alert" + aria-live="assertive" for error
 *   - aria-atomic="true" so screen readers read the whole pill on
 *     change (not just the delta)
 *
 * Dismiss is wired to the close button + an `Esc` keystroke when the
 * close button has focus. We don't add a document-level Esc trap
 * because that conflicts with dialogs.
 */
export function ToastContainer() {
  const toasts = useToastList();
  const { dismiss } = useToast();

  if (toasts.length === 0) return null;

  return (
    <div className="toast-stack" data-testid="toast-stack" aria-label="Notifications">
      {toasts.map((t) => {
        const isError = t.kind === 'error';
        return (
          <div
            key={t.id}
            className={`toast toast--${t.kind}`}
            data-testid={`toast-${t.kind}`}
            data-toast-id={t.id}
            role={isError ? 'alert' : 'status'}
            aria-live={isError ? 'assertive' : 'polite'}
            aria-atomic="true"
          >
            <span className="toast__icon" aria-hidden="true">
              {t.kind === 'success' ? '✓' : t.kind === 'error' ? '!' : 'ⓘ'}
            </span>
            <span className="toast__message">{t.message}</span>
            {t.action && (
              <button
                type="button"
                className="toast__action"
                data-testid={`toast-${t.kind}-action`}
                onClick={() => {
                  t.action!.onClick();
                  dismiss(t.id);
                }}
              >
                {t.action.label}
              </button>
            )}
            <button
              type="button"
              className="toast__close"
              data-testid={`toast-${t.kind}-close`}
              aria-label="Dismiss"
              onClick={() => dismiss(t.id)}
            >
              ×
            </button>
          </div>
        );
      })}
    </div>
  );
}
