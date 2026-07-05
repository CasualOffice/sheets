/*
 * Copyright (c) 2026 Casual Office. All rights reserved.
 */

/**
 * Streamable-HTTP transport for the MCP client — the standard way a browser
 * reaches a remote MCP server. Each JSON-RPC message is POSTed to the server
 * endpoint; the response body carries the matching JSON-RPC reply, which is
 * handed back to RpcConnection. Notifications (no id) POST and ignore the reply.
 *
 * This keeps McpClient transport-agnostic: swap this for a stdio bridge on the
 * desktop without touching the client or the agent.
 */

import type { JsonRpcTransport } from './jsonrpc';

export interface HttpMcpTransportOptions {
  /** Extra headers, e.g. Authorization for an authenticated MCP server. */
  headers?: Record<string, string>;
  /** Injected fetch (defaults to global). Lets tests supply their own. */
  fetchImpl?: typeof fetch;
}

export class HttpMcpTransport implements JsonRpcTransport {
  private handler: ((message: string) => void) | null = null;
  private readonly headers: Record<string, string>;
  private readonly doFetch: typeof fetch;
  private closed = false;

  constructor(
    private readonly url: string,
    options: HttpMcpTransportOptions = {},
  ) {
    this.headers = options.headers ?? {};
    // Bind to globalThis: native `fetch` throws "Illegal invocation" when
    // called as a method of another object (this === transport).
    this.doFetch = options.fetchImpl ?? fetch.bind(globalThis);
  }

  send(message: string): void {
    if (this.closed) return;
    void this.doFetch(this.url, {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        accept: 'application/json, text/event-stream',
        ...this.headers,
      },
      body: message,
    })
      .then(async (resp) => {
        const text = await resp.text();
        if (this.closed) return;
        // Accept both a bare JSON body and a single SSE `data:` frame.
        const payload = extractJson(text);
        if (payload) this.handler?.(payload);
      })
      .catch(() => {
        /* RpcConnection times out the pending request on its own. */
      });
  }

  onMessage(handler: (message: string) => void): void {
    this.handler = handler;
  }

  close(): void {
    this.closed = true;
  }
}

/** Pull the JSON-RPC object out of a plain body or an SSE `data:` frame. */
function extractJson(body: string): string | null {
  const trimmed = body.trim();
  if (!trimmed) return null;
  if (trimmed.startsWith('{') || trimmed.startsWith('[')) return trimmed;
  // SSE: take the last non-empty `data:` line.
  const dataLines = trimmed
    .split('\n')
    .map((l) => l.trim())
    .filter((l) => l.startsWith('data:'))
    .map((l) => l.slice(5).trim())
    .filter(Boolean);
  return dataLines.length ? dataLines[dataLines.length - 1] : null;
}
