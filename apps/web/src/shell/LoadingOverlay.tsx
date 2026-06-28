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
import { useLoading, type LoadingPhase } from '../loading-context';
import { Icon } from './Icon';
import { SOFT_WARN_BYTES } from './file-actions';
import { humanizeOpenError } from './humanize-error';

/**
 * Centered modal shown while a multi-MB workbook is being opened. The
 * overlay is intentionally render-blocking: clicks during the unit swap
 * race Univer's render phase and can leave the grid blank.
 *
 * Phase order is reading → parsing → mounting. We label each one
 * specifically so a 10 s wait *feels* like progress instead of a
 * mystery hang. The elapsed counter appears at 1.5 s — early enough
 * to reassure but late enough that fast opens never see it.
 */

const PHASE_TEXT: Record<LoadingPhase, string> = {
  reading: 'Reading file…',
  parsing: 'Parsing workbook…',
  mounting: 'Loading into the editor…',
};

/** Step number out of 3 for each phase — drives the "Step 2 of 3"
 *  indicator. Multi-second waits feel less like a hang when the user
 *  can see *which* step is happening and that progress is forward.
 *  We don't fake a progress bar (ExcelJS has no progress events) —
 *  just the discrete step count, which is honest. */
const PHASE_STEP: Record<LoadingPhase, number> = {
  reading: 1,
  parsing: 2,
  mounting: 3,
};
const TOTAL_STEPS = 3;

export function LoadingOverlay() {
  const { state, set } = useLoading();
  const [elapsed, setElapsed] = useState(0);

  useEffect(() => {
    if (!state || state.error) {
      setElapsed(0);
      return;
    }
    const id = setInterval(() => {
      setElapsed(Date.now() - state.startedAt);
    }, 250);
    return () => clearInterval(id);
  }, [state]);

  if (!state) return null;

  // Error flavor — same shell, different content. We keep the modal
  // blocking so the user explicitly dismisses; otherwise a fast-fail
  // would flash and vanish before they can read the message.
  //
  // The raw `state.error` is usually an ExcelJS / JSZip / parser
  // message that's gibberish to non-developers ("end of central
  // directory not found", "BIFF stream truncated", "HTTP 5xx"). Run
  // it through humanizeOpenError to get a friendly headline + one-
  // line hint; tuck the raw text into a collapsible <details> so
  // power users + bug reporters still get the diagnostic.
  if (state.error) {
    const retry = state.onRetry;
    const friendly = humanizeOpenError(state.error, state.fileName);
    return (
      <div
        className="loading-overlay"
        data-testid="loading-overlay"
        role="alertdialog"
        aria-live="assertive"
        aria-modal="true"
      >
        <div className="loading-overlay__card loading-overlay__card--error" data-testid="loading-overlay-error-card">
          <div className="loading-overlay__icon loading-overlay__icon--error" aria-hidden="true">
            <Icon name="error" size="md" />
          </div>
          <div className="loading-overlay__title">{friendly.title}</div>
          {friendly.hint && (
            <div
              className="loading-overlay__hint"
              data-testid="loading-overlay-hint"
              style={{
                marginTop: 8,
                fontSize: 13,
                lineHeight: 1.5,
                color: 'var(--color-text-secondary)',
                textAlign: 'center',
                maxWidth: 380,
              }}
            >
              {friendly.hint}
            </div>
          )}
          <details
            className="loading-overlay__details"
            data-testid="loading-overlay-error-details"
            style={{ marginTop: 14, fontSize: 12, color: 'var(--color-text-secondary)' }}
          >
            <summary style={{ cursor: 'pointer', userSelect: 'none' }}>
              Technical details
            </summary>
            <pre
              className="loading-overlay__error-text"
              data-testid="loading-overlay-error"
              style={{
                marginTop: 8,
                padding: 10,
                background: 'var(--color-surface-alt)',
                border: '1px solid var(--color-border)',
                borderRadius: 6,
                whiteSpace: 'pre-wrap',
                wordBreak: 'break-word',
                maxHeight: 140,
                overflowY: 'auto',
                textAlign: 'left',
              }}
            >
              {state.error}
            </pre>
          </details>
          <div className="loading-overlay__actions" style={{ marginTop: 16, display: 'flex', gap: 8, justifyContent: 'center' }}>
            <button
              type="button"
              className="btn-secondary"
              data-testid="loading-overlay-dismiss"
              onClick={() => set(null)}
            >
              Dismiss
            </button>
            {retry && (
              <button
                type="button"
                className="btn-primary"
                data-testid="loading-overlay-retry"
                onClick={() => {
                  set(null);
                  retry();
                }}
              >
                {friendly.retryLabel ?? 'Try again'}
              </button>
            )}
          </div>
        </div>
      </div>
    );
  }

  const showElapsed = elapsed > 1500;
  // Show the "big file" hint immediately once the user picks something
  // over the soft warning, instead of waiting 4 s — they already know
  // it's going to be slow; tell them up front.
  const isHugeFile = (state.sizeBytes ?? 0) > SOFT_WARN_BYTES;
  const showBigFileHint = isHugeFile || elapsed > 4000;
  const sizeText = state.sizeBytes ? formatBytes(state.sizeBytes) : null;

  return (
    <div className="loading-overlay" data-testid="loading-overlay" role="status" aria-live="polite">
      <div className="loading-overlay__card" data-testid="loading-overlay-card">
        <div className="loading-overlay__icon" aria-hidden="true">
          <div className="loading-overlay__icon-inner" />
        </div>
        <div className="loading-overlay__title" data-testid="loading-overlay-title">
          Opening <strong>{state.fileName}</strong>
          {sizeText && <span className="loading-overlay__size"> · {sizeText}</span>}
        </div>
        <div
          className="loading-overlay__step"
          data-testid="loading-overlay-step"
          aria-label={`Step ${PHASE_STEP[state.phase]} of ${TOTAL_STEPS}`}
        >
          Step {PHASE_STEP[state.phase]} of {TOTAL_STEPS}
        </div>
        <div className="loading-overlay__phase" data-testid="loading-overlay-phase">
          {PHASE_TEXT[state.phase]}
        </div>
        <div className="loading-overlay__bar" aria-hidden="true">
          <span className="loading-overlay__bar-fill" />
        </div>
        {showElapsed && (
          <div className="loading-overlay__elapsed" data-testid="loading-overlay-elapsed">
            {formatElapsed(elapsed)}
          </div>
        )}
        {showBigFileHint && (
          <div className="loading-overlay__hint">
            {isHugeFile
              ? 'This is a large workbook. Opening can take 10+ seconds and may use multiple GB of memory.'
              : 'Large workbooks can take a few extra seconds to open.'}
          </div>
        )}
      </div>
    </div>
  );
}

function formatBytes(b: number): string {
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(0)} KB`;
  const mb = b / (1024 * 1024);
  return mb < 10 ? `${mb.toFixed(1)} MB` : `${mb.toFixed(0)} MB`;
}

function formatElapsed(ms: number): string {
  const s = ms / 1000;
  if (s < 10) return `${s.toFixed(1)} s elapsed`;
  return `${Math.round(s)} s elapsed`;
}
