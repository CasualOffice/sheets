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
 * LLM transport for the Sheets AI panel.
 *
 * Mirrors the document editor's transport pattern (packages/react/src/docops/transport.ts):
 *  - DirectTransport  — browser fetch straight to Anthropic
 *  - CollabTransport  — WebSocket to the collab server /api/ai; server holds the
 *                       full LLM tool loop and routes tool_call messages back to
 *                       the originating client
 *
 *  - DesktopTransport — routes to the native llama.cpp model loaded in the
 *                       Casual Office desktop shell via the shared
 *                       `docops_llm_call` Tauri command (no key, no server).
 */

import { windowStringGlobal, viteEnv } from '../univer-facade';
import { isDesktop } from '../desk-bridge-bootstrap';

// ── Types ──────────────────────────────────────────────────────────────────

export type ToolExecutor = (toolName: string, args: Record<string, unknown>) => Promise<unknown>;

export interface LlmCallPayload {
  model: string;
  system: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  messages: any[];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  tools: any[];
  max_tokens: number;
  apiKey?: string;
  toolExecutor?: ToolExecutor;
  onText?: (text: string) => void;
  signal?: AbortSignal;
  maxToolRounds?: number;
}

export interface LlmCallResult {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  data: any;
  status: number;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  updatedHistory?: any[];
  capHit?: boolean;
}

export interface SheetsTransport {
  call(payload: LlmCallPayload): Promise<LlmCallResult>;
  readonly requiresApiKey: boolean;
  readonly drivesLoop: boolean;
}

// ── DirectTransport ────────────────────────────────────────────────────────

export class DirectTransport implements SheetsTransport {
  readonly requiresApiKey = true;
  readonly drivesLoop = false;

  async call(payload: LlmCallPayload): Promise<LlmCallResult> {
    if (!payload.apiKey) {
      return { data: { error: { message: 'No API key configured.' } }, status: 401 };
    }

    const useStream = !!payload.onText;

    const resp = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': payload.apiKey,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json',
      },
      body: JSON.stringify({
        model: payload.model,
        max_tokens: payload.max_tokens,
        system: payload.system,
        messages: payload.messages,
        tools: payload.tools,
        ...(useStream ? { stream: true } : {}),
      }),
      signal: payload.signal,
    });

    if (!useStream || !resp.body) {
      return { data: await resp.json(), status: resp.status };
    }

    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let buf = '';

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const content: any[] = [];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let msgDelta: any = {};

    try {
      outer: while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buf += decoder.decode(value, { stream: true });
        const lines = buf.split('\n');
        buf = lines.pop() ?? '';
        for (const line of lines) {
          if (!line.startsWith('data: ')) continue;
          const raw = line.slice(6).trim();
          if (raw === '[DONE]') break outer;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          let ev: any;
          try { ev = JSON.parse(raw); } catch { continue; }

          if (ev.type === 'content_block_start' && ev.content_block?.type === 'text') {
            content.push({ type: 'text', text: '' });
          } else if (ev.type === 'content_block_start' && ev.content_block?.type === 'tool_use') {
            content.push({ type: 'tool_use', id: ev.content_block.id, name: ev.content_block.name, input: {} });
          } else if (ev.type === 'content_block_delta') {
            const last = content[content.length - 1];
            if (ev.delta?.type === 'text_delta' && last?.type === 'text') {
              last.text += ev.delta.text ?? '';
              payload.onText?.(ev.delta.text ?? '');
            } else if (ev.delta?.type === 'input_json_delta' && last?.type === 'tool_use') {
              last._inputStr = (last._inputStr ?? '') + (ev.delta.partial_json ?? '');
            }
          } else if (ev.type === 'content_block_stop') {
            const last = content[content.length - 1];
            if (last?.type === 'tool_use' && last._inputStr) {
              try { last.input = JSON.parse(last._inputStr); } catch { /* leave empty */ }
              delete last._inputStr;
            }
          } else if (ev.type === 'message_delta') {
            msgDelta = ev.delta ?? {};
          }
        }
      }
    } catch (err) {
      return { data: { error: { message: String(err) } }, status: 500 };
    }

    return {
      data: { content, stop_reason: msgDelta.stop_reason ?? 'end_turn' },
      status: resp.status,
    };
  }
}

// ── CollabTransport ────────────────────────────────────────────────────────

/**
 * Routes AI orchestration through the collab server's /api/ai WebSocket.
 * The server holds the full LLM tool loop; tool_call messages are routed
 * back to this client which executes them via SheetsBridge.
 */
export class CollabTransport implements SheetsTransport {
  readonly requiresApiKey = false;
  readonly drivesLoop = true;

  constructor(private readonly aiWsUrl: string) {}

  call(payload: LlmCallPayload): Promise<LlmCallResult> {
    return new Promise((resolve, reject) => {
      if (payload.signal?.aborted) {
        reject(Object.assign(new Error('AbortError'), { name: 'AbortError' }));
        return;
      }

      let ws: WebSocket;
      try {
        ws = new WebSocket(this.aiWsUrl);
      } catch (err) {
        reject(new Error(`Failed to open AI WebSocket: ${String(err)}`));
        return;
      }

      let settled = false;

      const settle = (v: LlmCallResult | null, err?: Error) => {
        if (settled) return;
        settled = true;
        payload.signal?.removeEventListener('abort', onAbort);
        if (err) reject(err);
        else resolve(v!);
      };

      const onAbort = () => {
        try {
          ws.close(1000, 'aborted');
        } catch {
          /* ignore */
        }
        settle(null, Object.assign(new Error('AbortError'), { name: 'AbortError' }));
      };
      payload.signal?.addEventListener('abort', onAbort);

      ws.addEventListener('open', () => {
        ws.send(
          JSON.stringify({
            type: 'chat',
            model: payload.model,
            max_tokens: payload.max_tokens,
            system: payload.system,
            messages: payload.messages,
            tools: payload.tools,
            ...(payload.apiKey ? { apiKey: payload.apiKey } : {}),
            ...(payload.maxToolRounds != null ? { maxToolRounds: payload.maxToolRounds } : {}),
          }),
        );
      });

      ws.addEventListener('message', ({ data }: MessageEvent<string>) => {
        let msg: Record<string, unknown>;
        try {
          msg = JSON.parse(data) as Record<string, unknown>;
        } catch {
          settle(null, new Error('AI WS: received non-JSON frame'));
          ws.close();
          return;
        }

        if (msg.type === 'text') {
          payload.onText?.(msg.text as string);
        } else if (msg.type === 'tool_call') {
          const id = msg.id as string;
          const toolName = msg.toolName as string;
          const args = (msg.args ?? {}) as Record<string, unknown>;

          if (!payload.toolExecutor) {
            ws.send(
              JSON.stringify({
                type: 'tool_result',
                id,
                error: 'no toolExecutor configured on this client',
              }),
            );
            return;
          }

          payload
            .toolExecutor(toolName, args)
            .then((result) => {
              ws.send(JSON.stringify({ type: 'tool_result', id, result }));
            })
            .catch((err) => {
              ws.send(
                JSON.stringify({
                  type: 'tool_result',
                  id,
                  error: err instanceof Error ? err.message : String(err),
                }),
              );
            });
        } else if (msg.type === 'done') {
          settle({
            data: { ok: true },
            status: 200,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            updatedHistory: msg.history as any[],
            capHit: msg.capHit === true,
          });
        } else if (msg.type === 'error') {
          settle({
            data: { error: { message: msg.message as string } },
            status: 500,
          });
        }
      });

      ws.addEventListener('error', () => {
        settle(null, new Error('AI WebSocket connection failed'));
      });

      ws.addEventListener('close', ({ code, reason }: CloseEvent) => {
        if (!settled) {
          if (code === 1000 || reason === 'aborted') return;
          settle(null, new Error(`AI WebSocket closed unexpectedly (${code})`));
        }
      });
    });
  }
}

// ── DesktopTransport ─────────────────────────────────────────────────────────

/**
 * Routes AI calls to the native model loaded in the Casual Office desktop shell
 * (the llama.cpp `ai-worker`) through the shared `docops_llm_call` Tauri command
 * — the same backend the document editor uses. No API key, no collab server.
 *
 * drivesLoop=false: the panel drives the tool loop; each call() performs one LLM
 * round and returns an Anthropic-shaped `{ content, stop_reason }` response.
 */
export class DesktopTransport implements SheetsTransport {
  readonly requiresApiKey = false;
  readonly drivesLoop = false;

  async call(payload: LlmCallPayload): Promise<LlmCallResult> {
    const invoke = (
      window as unknown as {
        __TAURI__?: { core?: { invoke?: (cmd: string, args?: unknown) => Promise<unknown> } };
      }
    ).__TAURI__?.core?.invoke;
    if (!invoke) {
      return {
        data: { error: { message: 'Native AI is only available in the desktop app.' } },
        status: 500,
      };
    }
    try {
      // The Rust command is docops_llm_call(args: DocopsLlmArgs), so the payload
      // MUST be nested under `args` (a bare object throws "missing required key
      // args"). Inner fields are camelCase (#[serde(rename_all = "camelCase")]).
      const data = await invoke('docops_llm_call', {
        args: {
          model: payload.model,
          system: payload.system,
          messages: payload.messages,
          tools: payload.tools,
          maxTokens: payload.max_tokens,
          apiKey: payload.apiKey ?? '',
        },
      });
      return { data, status: 200 };
    } catch (err) {
      return { data: { error: { message: String(err) } }, status: 500 };
    }
  }
}

// ── Factory ────────────────────────────────────────────────────────────────

/**
 * Derives the AI WebSocket URL from the collab server endpoint.
 * Replaces /yjs with /api/ai — same base as Yjs but the AI route.
 */
function sheetsAiWsUrl(): string {
  const base =
    windowStringGlobal('__COLLAB_WS_URL__') ??
    viteEnv('VITE_COLLAB_WS_URL') ??
    `${window.location.protocol === 'https:' ? 'wss:' : 'ws:'}//${window.location.host}/yjs`;
  return base.replace(/\/yjs$/, '').replace(/\/+$/, '') + '/api/ai';
}

/**
 * Returns true when a collab server WS URL is configured.
 * Used by the UI to gate the AI panel button — on plain web without a collab
 * server, the DirectTransport API-key flow is not yet exposed to end users.
 */
export function hasCollabServer(): boolean {
  return !!windowStringGlobal('__COLLAB_WS_URL__') || !!viteEnv('VITE_COLLAB_WS_URL');
}

/**
 * Explicit opt-in to expose the AI panel on plain web (DirectTransport,
 * bring-your-own Anthropic key). Off by default so the key form isn't shown
 * to end users; a self-hoster sets `window.__ENABLE_AI__ = true` or builds
 * with `VITE_ENABLE_AI=1`. The e2e suite sets the window global.
 */
export function aiUiForced(): boolean {
  return windowStringGlobal('__ENABLE_AI__') === 'true' || !!viteEnv('VITE_ENABLE_AI');
}

/**
 * Returns the appropriate transport for the current environment:
 *  - CollabTransport when the collab server is available
 *  - DirectTransport otherwise (user provides Anthropic key)
 */
export function createSheetsTransport(): SheetsTransport {
  // Desktop shell: use the loaded native model. Its own bridge (?desk=1) means
  // no collab server and no user API key are needed.
  if (isDesktop()) {
    return new DesktopTransport();
  }
  if (hasCollabServer()) {
    return new CollabTransport(sheetsAiWsUrl());
  }
  return new DirectTransport();
}
