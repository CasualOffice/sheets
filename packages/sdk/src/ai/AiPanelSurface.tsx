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
 * AI surface for the Sheets SDK.
 *
 * `<CasualSheets ai={…}>` opts a host into a supported AI task-pane surface,
 * mirroring how the docs `CasualEditor`/`DocxEditor` SDK exposes its DocOps
 * panel behind a `docopsTransport` prop. The SDK owns the prop contract, the
 * `SheetsAiTransport` type (see `./transport`), and the layout slot beside the
 * grid; the host supplies the panel body via `ai.render` and drives its tool
 * loop against the transport.
 *
 * MVP (slot + transport wiring). The reference implementation of the panel
 * body — chat UI, `SheetsBridge` (FUniver tool catalog), agent mode, MCP — is
 * still in the reference app at `apps/web/src/shell/AiPanel.tsx`; that panel is
 * entangled with app-only modules (`use-ui`, `Icon`, the `ai/` catalog + MCP
 * runtime), so it is NOT yet moved wholesale into the SDK.
 *
 * TODO(sheets#280): extract the `SheetsBridge` catalog + a self-contained
 * default panel body into the SDK so `ai.enabled` renders a working assistant
 * with no `render` slot required.
 */

import type { ReactNode } from 'react';
import type { CasualSheetsAPI } from '../sheets/api';
import type { SheetsAiTransport } from './transport';

/** Lifecycle / progress signals a panel body reports back to the host. */
export type SheetsAiAction =
  | { type: 'open' }
  | { type: 'close' }
  | { type: 'message'; role: 'user' | 'assistant'; text: string }
  | { type: 'tool'; tool: string; status: 'running' | 'done' | 'error' }
  | { type: 'error'; message: string };

/** Context handed to `ai.render` — everything a panel body needs to run. */
export interface SheetsAiRenderContext {
  /** The ready editor facade (`CasualSheetsAPI`). Never null inside `render`. */
  api: CasualSheetsAPI;
  /** The transport the host configured (if any). */
  transport?: SheetsAiTransport;
  /** Forward panel lifecycle/progress to the host's `ai.onAction`. */
  onAction?: (action: SheetsAiAction) => void;
  /** Request the surface be dismissed (emits an `onAction({type:'close'})`). */
  close: () => void;
}

/** The `ai` prop on `<CasualSheets>`. */
export interface SheetsAiConfig {
  /** Mount the AI surface beside the grid. Default `false` (no surface). */
  enabled?: boolean;
  /** LLM transport the panel body drives its tool loop against. Build one with
   *  `createSheetsAiTransport()` or supply your own `SheetsAiTransport`. */
  transport?: SheetsAiTransport;
  /** Observe panel lifecycle + progress (open/close, messages, tool steps). */
  onAction?: (action: SheetsAiAction) => void;
  /** Render the panel body. Receives a `SheetsAiRenderContext` (ready `api` +
   *  `transport` + `onAction` + `close`). Until the SDK ships a built-in body,
   *  this slot is how the panel is mounted — the reference app passes its
   *  `<AiPanel>` here. When omitted, `enabled` mounts an empty pane. */
  render?: (ctx: SheetsAiRenderContext) => ReactNode;
}

const asideStyle: React.CSSProperties = {
  flex: '0 0 auto',
  width: 340,
  maxWidth: '100%',
  height: '100%',
  minHeight: 0,
  display: 'flex',
  flexDirection: 'column',
  borderLeft: '1px solid var(--color-divider, #e5e7eb)',
  background: 'var(--color-surface, #ffffff)',
  overflow: 'hidden',
};

/**
 * Renders the configured AI surface beside the grid. Returns `null` until the
 * editor is ready (`api` populated) or when disabled — so bare-grid consumers
 * pay nothing. The host's `render` slot owns the panel body; if absent, an
 * empty (but present) pane is mounted so layout is stable.
 */
export function AiPanelSurface({
  config,
  api,
}: {
  config: SheetsAiConfig | undefined;
  api: CasualSheetsAPI | null;
}): ReactNode {
  if (!config?.enabled || !api) return null;

  const ctx: SheetsAiRenderContext = {
    api,
    transport: config.transport,
    onAction: config.onAction,
    close: () => config.onAction?.({ type: 'close' }),
  };

  return (
    <aside data-testid="casual-sheets-ai" style={asideStyle}>
      {config.render?.(ctx) ?? null}
    </aside>
  );
}
