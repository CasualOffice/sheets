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

import { useState } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { Dialog } from './Dialog';
import { Icon } from './Icon';
import {
  applyWatermark,
  DEFAULT_WATERMARK_OPACITY,
  DEFAULT_WATERMARK_TEXT,
  type WatermarkConfig,
} from './watermark';

/**
 * Watermark config dialog opened from View → Confidential watermark.
 *
 * Replaces the old hardcoded on/off toggle with a small surface:
 *   - On/off switch (drives apply vs. clear).
 *   - Custom text (defaults to CONFIDENTIAL; free text).
 *   - Opacity (a coarse select — the service supports any 0–1 value but
 *     a few stops cover the practical range without a fiddly slider).
 *
 * Apply / clear is delegated to `watermark.ts`, which talks to Univer's
 * WatermarkService and persists the chosen text/opacity to localStorage.
 */

type Props = {
  api: FUniver;
  /** Whether a watermark is currently applied (seeds the switch). */
  initialOn: boolean;
  /** Last chosen / currently applied text + opacity (seeds the fields). */
  initial: WatermarkConfig;
  onClose: () => void;
  /** Called after apply/clear so the menu's checked state stays in sync. */
  onApplied: (on: boolean) => void;
};

const OPACITY_STOPS: { label: string; value: number }[] = [
  { label: 'Light (8%)', value: 0.08 },
  {
    label: `Default (${Math.round(DEFAULT_WATERMARK_OPACITY * 100)}%)`,
    value: DEFAULT_WATERMARK_OPACITY,
  },
  { label: 'Medium (20%)', value: 0.2 },
  { label: 'Strong (35%)', value: 0.35 },
];

export function WatermarkDialog({ api, initialOn, initial, onClose, onApplied }: Props) {
  const [on, setOn] = useState(initialOn);
  const [text, setText] = useState(initial.text);
  const [opacity, setOpacity] = useState(initial.opacity);

  const apply = () => {
    applyWatermark(api, on, { text, opacity });
    onApplied(on);
    onClose();
  };

  // `on` gates the text/opacity controls — off means there's nothing to
  // configure, matching Excel/Word where the watermark options grey out
  // when the watermark is removed.
  const disabled = !on;

  return (
    <Dialog
      title="Watermark"
      onClose={onClose}
      data-testid="watermark-dialog"
      footer={
        <>
          <button
            type="button"
            className="btn-secondary"
            data-testid="watermark-cancel"
            onClick={onClose}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="watermark-apply"
            onClick={apply}
          >
            Apply
          </button>
        </>
      }
    >
      <div className="watermark">
        <label className="watermark__switch">
          <span className="watermark__switch-label">
            <Icon name="water_drop" size="sm" />
            Show watermark
          </span>
          <input
            type="checkbox"
            role="switch"
            data-testid="watermark-toggle"
            checked={on}
            onChange={(e) => setOn(e.target.checked)}
          />
        </label>

        <label className="watermark__field">
          <span>Text</span>
          <input
            type="text"
            data-testid="watermark-text"
            value={text}
            onChange={(e) => setText(e.target.value)}
            placeholder={DEFAULT_WATERMARK_TEXT}
            disabled={disabled}
            maxLength={60}
            spellCheck={false}
          />
        </label>

        <label className="watermark__field">
          <span>Opacity</span>
          <select
            className="watermark__select"
            data-testid="watermark-opacity"
            value={String(opacity)}
            onChange={(e) => setOpacity(Number(e.target.value))}
            disabled={disabled}
          >
            {OPACITY_STOPS.map((stop) => (
              <option key={stop.value} value={String(stop.value)}>
                {stop.label}
              </option>
            ))}
          </select>
        </label>

        <p className="watermark__note">
          The watermark is shown on screen only and is not saved into exported files.
        </p>
      </div>
    </Dialog>
  );
}
