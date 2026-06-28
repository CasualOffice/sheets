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

import { useEffect, useState } from 'react';
import { useSaveStatus } from './save-status-context';

/**
 * Title-bar pill that surfaces the save status. UX_AUDIT.md §4.3.
 *
 * States:
 *   idle    — render nothing (no clutter when there's nothing to say)
 *   saving  — "Saving…" with a spinner
 *   saved   — "Saved" + relative-time ("just now" / "2 min ago")
 *   error   — "Save failed" with the underlying message as a tooltip
 *
 * The relative-time string ticks every 30 s so it stays roughly
 * correct without burning a render every second. Cleared whenever
 * the status moves off `saved`.
 */
export function SaveStatusPill() {
  const { status } = useSaveStatus();

  const [, forceTick] = useState(0);
  useEffect(() => {
    if (status.kind !== 'saved') return;
    const id = setInterval(() => forceTick((n) => n + 1), 30_000);
    return () => clearInterval(id);
  }, [status.kind]);

  if (status.kind === 'idle') return null;

  let label = '';
  let tone: 'neutral' | 'success' | 'error' = 'neutral';
  let title = '';
  if (status.kind === 'saving') {
    label = 'Saving…';
    tone = 'neutral';
  } else if (status.kind === 'saved') {
    label = `Saved ${formatRelative(Date.now() - status.savedAt)}`;
    tone = 'success';
    title = `Saved at ${new Date(status.savedAt).toLocaleTimeString()}`;
  } else if (status.kind === 'error') {
    label = 'Save failed';
    tone = 'error';
    title = status.message;
  }

  // Design-system Pill tones (height 22 / 0 9px / pill radius), tokenised so
  // they auto-theme — matches the DS Pill used elsewhere.
  const TONE: Record<typeof tone, { bg: string; fg: string; bd: string }> = {
    success: {
      bg: 'var(--color-success-soft)',
      fg: 'var(--color-success)',
      bd: 'color-mix(in srgb, var(--color-success) 30%, transparent)',
    },
    error: {
      bg: 'var(--color-danger-soft)',
      fg: 'var(--color-danger)',
      bd: 'color-mix(in srgb, var(--color-danger) 30%, transparent)',
    },
    neutral: {
      bg: 'var(--color-toolbar-pill)',
      fg: 'var(--color-text-secondary)',
      bd: 'var(--color-border)',
    },
  };

  return (
    <span
      className={`save-status-pill save-status-pill--${tone}`}
      data-testid="save-status-pill"
      data-state={status.kind}
      role="status"
      aria-live="polite"
      title={title}
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        gap: 5,
        height: 22,
        padding: '0 9px',
        borderRadius: 'var(--radius-pill)',
        fontSize: 'var(--text-sm)',
        fontWeight: 'var(--weight-medium)',
        background: TONE[tone].bg,
        color: TONE[tone].fg,
        border: `1px solid ${TONE[tone].bd}`,
      }}
    >
      {status.kind === 'saving' && (
        <span
          aria-hidden="true"
          style={{
            width: 10,
            height: 10,
            border: '2px solid currentColor',
            borderRightColor: 'transparent',
            borderRadius: '50%',
            animation: 'save-status-spin 800ms linear infinite',
          }}
        />
      )}
      <span>{label}</span>
    </span>
  );
}

function formatRelative(ms: number): string {
  if (ms < 0 || !Number.isFinite(ms)) return 'just now';
  const sec = Math.round(ms / 1000);
  if (sec < 30) return 'just now';
  if (sec < 90) return '1 min ago';
  const min = Math.round(sec / 60);
  if (min < 60) return `${min} min ago`;
  const hr = Math.round(min / 60);
  if (hr < 24) return `${hr} hr ago`;
  const day = Math.round(hr / 24);
  return `${day} day${day === 1 ? '' : 's'} ago`;
}
