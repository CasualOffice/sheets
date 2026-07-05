/*
 * Copyright (c) 2026 Casual Office. All rights reserved.
 */

/**
 * Minimal JSON-RPC 2.0 client for the Model Context Protocol.
 *
 * Framing is injected as a `JsonRpcTransport` so the same connection works over
 * a desktop stdio pipe, a WebSocket to the collab server, or an in-memory pair
 * in tests. This layer owns id assignment, request/response correlation, and
 * notification dispatch — nothing MCP-specific lives here.
 */

export interface JsonRpcTransport {
  /** Send one framed JSON-RPC message (already serialized). */
  send(message: string): void;
  /** Register the handler that receives incoming framed messages. */
  onMessage(handler: (message: string) => void): void;
  close?(): void;
}

interface Pending {
  resolve: (result: unknown) => void;
  reject: (error: Error) => void;
}

export class RpcConnection {
  private nextId = 1;
  private readonly pending = new Map<number, Pending>();
  private readonly notify = new Map<string, (params: unknown) => void>();

  constructor(private readonly transport: JsonRpcTransport) {
    transport.onMessage((raw) => this.handle(raw));
  }

  /** Send a request and await its result. */
  request<T = unknown>(method: string, params?: unknown, timeoutMs = 20000): Promise<T> {
    const id = this.nextId++;
    const payload = JSON.stringify({ jsonrpc: '2.0', id, method, params });
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(() => {
        if (this.pending.delete(id)) {
          reject(new Error(`MCP request '${method}' timed out after ${timeoutMs}ms`));
        }
      }, timeoutMs);
      this.pending.set(id, {
        resolve: (r) => {
          clearTimeout(timer);
          resolve(r as T);
        },
        reject: (e) => {
          clearTimeout(timer);
          reject(e);
        },
      });
      try {
        this.transport.send(payload);
      } catch (err) {
        this.pending.delete(id);
        clearTimeout(timer);
        reject(err instanceof Error ? err : new Error(String(err)));
      }
    });
  }

  /** Fire-and-forget notification (no id, no response). */
  notifyServer(method: string, params?: unknown): void {
    this.transport.send(JSON.stringify({ jsonrpc: '2.0', method, params }));
  }

  /** Subscribe to a server-initiated notification method. */
  onNotification(method: string, handler: (params: unknown) => void): void {
    this.notify.set(method, handler);
  }

  close(): void {
    for (const [, p] of this.pending) p.reject(new Error('MCP connection closed'));
    this.pending.clear();
    this.transport.close?.();
  }

  private handle(raw: string): void {
    let msg: {
      id?: number;
      method?: string;
      params?: unknown;
      result?: unknown;
      error?: { code: number; message: string };
    };
    try {
      msg = JSON.parse(raw);
    } catch {
      return; // ignore malformed frames
    }

    // Response to one of our requests.
    if (typeof msg.id === 'number' && (msg.result !== undefined || msg.error !== undefined)) {
      const p = this.pending.get(msg.id);
      if (!p) return;
      this.pending.delete(msg.id);
      if (msg.error) p.reject(new Error(`MCP error ${msg.error.code}: ${msg.error.message}`));
      else p.resolve(msg.result);
      return;
    }

    // Server notification.
    if (typeof msg.method === 'string' && msg.id === undefined) {
      this.notify.get(msg.method)?.(msg.params);
    }
  }
}
