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
 * SDK-owned LLM transport contract for the Sheets AI surface.
 *
 * This is the canonical `SheetsAiTransport` type that the `ai` prop on
 * `<CasualSheets>` accepts. It mirrors the reference app's transport
 * (`apps/web/src/ai/transport.ts`) but lives in the SDK so integrators can
 * type their transport straight off `@casualoffice/sheets` and reuse the
 * shipped implementations:
 *   - `DirectAiTransport`  — browser fetch straight to Anthropic (BYO key)
 *   - `CollabAiTransport`  — WebSocket to the collab server `/api/ai`; the
 *                            server holds the key and runs ONE LLM round per
 *                            `call()` (single-round), routing `tool_call`
 *                            frames back to this client's `toolExecutor`
 *   - `DesktopAiTransport` — native llama.cpp model in the Casual Office
 *                            desktop shell via the `docops_llm_call` Tauri
 *                            command (no key, no server)
 *
 * The panel drives the tool loop (`drivesLoop = false` on every transport):
 * each `call()` performs one round and returns an Anthropic-shaped
 * `{ content, stop_reason }` payload.
 */

// ── Types ──────────────────────────────────────────────────────────────────

export type SheetsAiToolExecutor = (
  toolName: string,
  args: Record<string, unknown>,
) => Promise<unknown>;

export interface SheetsAiLlmPayload {
  model: string;
  system: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  messages: any[];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  tools: any[];
  max_tokens: number;
  apiKey?: string;
  toolExecutor?: SheetsAiToolExecutor;
  onText?: (text: string) => void;
  signal?: AbortSignal;
  maxToolRounds?: number;
}

export interface SheetsAiLlmResult {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  data: any;
  status: number;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  updatedHistory?: any[];
  capHit?: boolean;
}

export interface SheetsAiTransport {
  call(payload: SheetsAiLlmPayload): Promise<SheetsAiLlmResult>;
  readonly requiresApiKey: boolean;
  readonly drivesLoop: boolean;
}

// ── DirectAiTransport ──────────────────────────────────────────────────────

export class DirectAiTransport implements SheetsAiTransport {
  readonly requiresApiKey = true;
  readonly drivesLoop = false;

  async call(payload: SheetsAiLlmPayload): Promise<SheetsAiLlmResult> {
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
          try {
            ev = JSON.parse(raw);
          } catch {
            continue;
          }

          if (ev.type === 'content_block_start' && ev.content_block?.type === 'text') {
            content.push({ type: 'text', text: '' });
          } else if (ev.type === 'content_block_start' && ev.content_block?.type === 'tool_use') {
            content.push({
              type: 'tool_use',
              id: ev.content_block.id,
              name: ev.content_block.name,
              input: {},
            });
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
              try {
                last.input = JSON.parse(last._inputStr);
              } catch {
                /* leave empty */
              }
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

// ── CollabAiTransport ──────────────────────────────────────────────────────

/**
 * Routes AI orchestration through the collab server's `/api/ai` WebSocket.
 * The server holds the Anthropic key; `singleRound: true` means the server
 * runs ONE LLM round per `call()` and routes any `tool_call` frames back to
 * this client, which executes them via the payload's `toolExecutor`. The
 * panel/agent drives the multi-round loop from there (so `drivesLoop = false`).
 */
export class CollabAiTransport implements SheetsAiTransport {
  readonly requiresApiKey = false;
  readonly drivesLoop = false;

  constructor(private readonly aiWsUrl: string) {}

  call(payload: SheetsAiLlmPayload): Promise<SheetsAiLlmResult> {
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

      const settle = (v: SheetsAiLlmResult | null, err?: Error) => {
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
            singleRound: true,
            ...(payload.apiKey ? { apiKey: payload.apiKey } : {}),
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

        if (msg.type === 'round') {
          // Single-round reply: the panel/agent drives the loop from here.
          settle({
            data: {
              content: msg.content ?? [],
              stop_reason: (msg.stop_reason as string) ?? 'end_turn',
            },
            status: 200,
          });
        } else if (msg.type === 'text') {
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

// ── DesktopAiTransport ─────────────────────────────────────────────────────

/**
 * Routes AI calls to the native model loaded in the Casual Office desktop
 * shell (the llama.cpp worker) through the shared `docops_llm_call` Tauri
 * command — the same backend the document editor uses. No API key, no collab
 * server. Each `call()` performs one LLM round.
 */
export class DesktopAiTransport implements SheetsAiTransport {
  readonly requiresApiKey = false;
  readonly drivesLoop = false;

  async call(payload: SheetsAiLlmPayload): Promise<SheetsAiLlmResult> {
    const invoke = tauriInvoke();
    if (!invoke) {
      return {
        data: { error: { message: 'Native AI is only available in the desktop app.' } },
        status: 500,
      };
    }
    try {
      // The Rust command is docops_llm_call(args: DocopsLlmArgs), so the
      // payload MUST be nested under `args`. Inner fields are camelCase.
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

// ── Environment helpers ──────────────────────────────────────────────────────

/** Tauri `invoke` when running inside the Casual Office desktop shell, else null. */
function tauriInvoke(): ((cmd: string, args?: unknown) => Promise<unknown>) | null {
  const inv = (
    window as unknown as {
      __TAURI__?: { core?: { invoke?: (c: string, a?: unknown) => Promise<unknown> } };
    }
  ).__TAURI__?.core?.invoke;
  return typeof inv === 'function' ? inv : null;
}

/**
 * Derives the AI WebSocket URL: an explicit override wins, else the
 * `__COLLAB_WS_URL__` window global (same base as Yjs, `/yjs` → `/api/ai`),
 * else same-origin `/api/ai`.
 */
function resolveAiWsUrl(explicit?: string): string {
  const base =
    explicit ??
    (window as unknown as { __COLLAB_WS_URL__?: string }).__COLLAB_WS_URL__ ??
    `${window.location.protocol === 'https:' ? 'wss:' : 'ws:'}//${window.location.host}/yjs`;
  return base.replace(/\/yjs$/, '').replace(/\/+$/, '') + '/api/ai';
}

// ── Factory ──────────────────────────────────────────────────────────────────

export interface CreateSheetsAiTransportOptions {
  /** Explicit collab WS base URL (e.g. `wss://host/yjs`). Falls back to the
   *  `__COLLAB_WS_URL__` window global, then same-origin. */
  collabWsUrl?: string;
  /** Force the browser-direct Anthropic transport (bring-your-own key) even
   *  when a collab server is configured. */
  forceDirect?: boolean;
}

/**
 * Picks the transport for the current environment:
 *  - `DesktopAiTransport` inside the desktop shell (native model),
 *  - `CollabAiTransport` when a collab WS URL is available (server holds key),
 *  - `DirectAiTransport` otherwise (user provides an Anthropic key).
 */
export function createSheetsAiTransport(
  opts: CreateSheetsAiTransportOptions = {},
): SheetsAiTransport {
  if (!opts.forceDirect && tauriInvoke()) {
    return new DesktopAiTransport();
  }
  const hasCollab =
    !!opts.collabWsUrl || !!(window as unknown as { __COLLAB_WS_URL__?: string }).__COLLAB_WS_URL__;
  if (!opts.forceDirect && hasCollab) {
    return new CollabAiTransport(resolveAiWsUrl(opts.collabWsUrl));
  }
  return new DirectAiTransport();
}
