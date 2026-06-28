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
 * Shared inline styles for built-in SDK dialogs. Kept in their own module (not
 * in Dialog.tsx) so the component files stay component-only for React Fast
 * Refresh. Every value reads a `--cs-chrome-*` token with a hardcoded fallback,
 * matching the chrome's theming approach (light/dark via the `data-theme`
 * wrapper `<CasualSheets>` sets).
 */

import type { CSSProperties } from 'react';

export const DIALOG_FIELD_STYLE: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: 4,
  marginBottom: 12,
};

export const DIALOG_LABEL_STYLE: CSSProperties = {
  fontSize: 12,
  fontWeight: 500,
  color: 'var(--cs-chrome-muted, #605e5c)',
};

export const DIALOG_INPUT_STYLE: CSSProperties = {
  height: 30,
  padding: '0 8px',
  border: '1px solid var(--cs-chrome-border, #cdd3db)',
  borderRadius: 6,
  background: 'var(--cs-chrome-input-bg, #fff)',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
  boxSizing: 'border-box',
};

export const DIALOG_BTN_PRIMARY_STYLE: CSSProperties = {
  height: 30,
  padding: '0 14px',
  border: 'none',
  borderRadius: 6,
  background: 'var(--cs-chrome-active-fg, #0e7490)',
  color: '#fff',
  font: 'inherit',
  fontSize: 13,
  fontWeight: 500,
  cursor: 'pointer',
};

export const DIALOG_BTN_SECONDARY_STYLE: CSSProperties = {
  height: 30,
  padding: '0 14px',
  border: '1px solid var(--cs-chrome-border, #cdd3db)',
  borderRadius: 6,
  background: 'var(--cs-chrome-input-bg, #fff)',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
  cursor: 'pointer',
};
