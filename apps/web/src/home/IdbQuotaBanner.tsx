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
import { Icon } from '../shell/Icon';

/**
 * Mode 1 (browser-only) soft warning when IDB use crosses a threshold.
 *
 * IndexedDB doesn't enforce a hard quota across browsers (~50–60% of
 * free disk, evicted under pressure). On the hosted demo / GitHub Pages
 * deploy that's where the user's recent files + autosave + version
 * history live; once they cross ~50 MB they're closer to eviction than
 * they realise.
 *
 * Banner copy intentionally non-alarming — eviction is opportunistic,
 * not imminent. The follow-up the user can act on is "Pin a folder"
 * (gets the FSA save bypass) or "Save your important files to disk"
 * before they're at risk.
 *
 * Hidden when:
 *   - `navigator.storage.estimate` isn't available (Safari old, Firefox
 *      private mode) → can't probe → don't speculate
 *   - usage < threshold
 *   - user already dismissed this session
 */

const THRESHOLD_BYTES = 50 * 1024 * 1024;
const DISMISS_KEY = 'casual-sheets:idb-quota-banner-dismissed';

export function IdbQuotaBanner() {
  const [usage, setUsage] = useState<number | null>(null);
  const [dismissed, setDismissed] = useState<boolean>(() => {
    try {
      return sessionStorage.getItem(DISMISS_KEY) === '1';
    } catch {
      return false;
    }
  });

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      const navStorage = (
        navigator as Navigator & {
          storage?: { estimate?: () => Promise<{ usage?: number; quota?: number }> };
        }
      ).storage;
      if (!navStorage?.estimate) return;
      try {
        const estimate = await navStorage.estimate();
        if (cancelled) return;
        setUsage(estimate.usage ?? 0);
      } catch {
        /* permission denied / private mode — silent fail */
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  if (dismissed || usage === null || usage < THRESHOLD_BYTES) return null;

  const dismiss = () => {
    try {
      sessionStorage.setItem(DISMISS_KEY, '1');
    } catch {
      /* sessionStorage blocked — banner still hides for this render */
    }
    setDismissed(true);
  };

  const usedMb = (usage / (1024 * 1024)).toFixed(0);

  return (
    <div className="home__reopen home__reopen--quota" role="status" data-testid="idb-quota-banner">
      <span className="home__reopen-icon" aria-hidden>
        <Icon name="storage" />
      </span>
      <span className="home__reopen-text">
        You’re using <strong>~{usedMb} MB</strong> of browser storage. Pin a folder or save
        important files to disk so you don’t lose them if your browser evicts the tab cache.
      </span>
      <button
        type="button"
        className="home__reopen-btn"
        onClick={dismiss}
        aria-label="Dismiss"
        data-testid="idb-quota-dismiss"
      >
        <Icon name="close" size="sm" />
      </button>
    </div>
  );
}
