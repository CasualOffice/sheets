/*
 * Copyright (c) 2026 Casual Office. All rights reserved.
 */

/**
 * Streamable-HTTP transport for the MCP client — the standard way a browser
 * reaches a remote MCP server (MCP spec 2025-06-18). Each JSON-RPC message is
 * POSTed to the server endpoint. The reply may come back as:
 *   - a buffered `application/json` body (one JSON-RPC object), or
 *   - a `text/event-stream` (SSE) body, possibly long-lived, delivering one or
 *     more `data:` JSON-RPC frames.
 *
 * We must NOT `await resp.text()` on the SSE case — that blocks until the stream
 * closes, hanging every request against a spec-compliant streaming server until
 * the RpcConnection timeout. Instead we read the body incrementally and dispatch
 * each SSE frame as it arrives.
 *
 * This keeps McpClient transport-agnostic: swap this for a stdio bridge on the
 * desktop without touching the client or the agent.
 */

import type { JsonRpcTransport } from './jsonrpc';

const DEFAULT_PROTOCOL_VERSION = '2025-06-18';

export interface HttpMcpTransportOptions {
  /** Extra headers, e.g. Authorization for an authenticated MCP server. */
  headers?: Record<string, string>;
  /** Injected fetch (defaults to global). Lets tests supply their own. */
  fetchImpl?: typeof fetch;
  /** MCP protocol version echoed on every request after initialize. */
  protocolVersion?: string;
}

export class HttpMcpTransport implements JsonRpcTransport {
  private handler: ((message: string) => void) | null = null;
  private readonly headers: Record<string, string>;
  private readonly doFetch: typeof fetch;
  private readonly protocolVersion: string;
  private readonly inflight = new Set<AbortController>();
  // Captured from the initialize response — stateful servers require it echoed
  // on every subsequent request (spec: Mcp-Session-Id).
  private sessionId: string | null = null;
  private closed = false;

  constructor(
    private readonly url: string,
    options: HttpMcpTransportOptions = {},
  ) {
    this.headers = options.headers ?? {};
    this.protocolVersion = options.protocolVersion ?? DEFAULT_PROTOCOL_VERSION;
    // Bind to globalThis: native `fetch` throws "Illegal invocation" when
    // called as a method of another object (this === transport).
    this.doFetch = options.fetchImpl ?? fetch.bind(globalThis);
  }

  send(message: string): void {
    if (this.closed) return;
    const controller = new AbortController();
    this.inflight.add(controller);
    void this.doFetch(this.url, {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        accept: 'application/json, text/event-stream',
        'mcp-protocol-version': this.protocolVersion,
        ...(this.sessionId ? { 'mcp-session-id': this.sessionId } : {}),
        ...this.headers,
      },
      body: message,
      signal: controller.signal,
    })
      .then(async (resp) => {
        const sid = resp.headers.get('mcp-session-id');
        if (sid) this.sessionId = sid;
        if (this.closed) return;
        const ctype = resp.headers.get('content-type') ?? '';
        if (ctype.includes('text/event-stream') && resp.body) {
          await this.readSse(resp.body);
        } else {
          const payload = extractJson(await resp.text());
          if (payload && !this.closed) this.handler?.(payload);
        }
      })
      .catch(() => {
        /* aborted, or RpcConnection times out the pending request on its own. */
      })
      .finally(() => this.inflight.delete(controller));
  }

  /** Read an SSE body incrementally, dispatching each `data:` frame as it lands. */
  private async readSse(body: ReadableStream<Uint8Array>): Promise<void> {
    const reader = body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    try {
      for (;;) {
        const { done, value } = await reader.read();
        if (done || this.closed) break;
        // Normalise CRLF so a single frame delimiter works.
        buffer += decoder.decode(value, { stream: true }).replace(/\r\n/g, '\n');
        let sep: number;
        while ((sep = buffer.indexOf('\n\n')) !== -1) {
          const frame = buffer.slice(0, sep);
          buffer = buffer.slice(sep + 2);
          const data = frameData(frame);
          if (data && !this.closed) this.handler?.(data);
        }
      }
    } catch {
      /* stream aborted on close */
    } finally {
      reader.releaseLock();
    }
  }

  onMessage(handler: (message: string) => void): void {
    this.handler = handler;
  }

  close(): void {
    this.closed = true;
    for (const c of this.inflight) c.abort();
    this.inflight.clear();
  }
}

/** Join the `data:` lines of a single SSE frame into one JSON-RPC payload. */
function frameData(frame: string): string | null {
  const data = frame
    .split('\n')
    .filter((l) => l.startsWith('data:'))
    .map((l) => l.slice(5).replace(/^ /, ''))
    .join('\n')
    .trim();
  return data || null;
}

/** Pull the JSON-RPC object out of a plain body or a buffered SSE body. */
function extractJson(body: string): string | null {
  const trimmed = body.trim();
  if (!trimmed) return null;
  if (trimmed.startsWith('{') || trimmed.startsWith('[')) return trimmed;
  const dataLines = trimmed
    .split('\n')
    .map((l) => l.trim())
    .filter((l) => l.startsWith('data:'))
    .map((l) => l.slice(5).trim())
    .filter(Boolean);
  return dataLines.length ? dataLines[dataLines.length - 1] : null;
}
