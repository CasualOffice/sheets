/*
 * Copyright (c) 2026 Casual Office. All rights reserved.
 */

/**
 * MCP server core: answers the Model Context Protocol over an injected
 * transport, exposing a set of tools (the DocOps catalog) to any MCP client —
 * Claude Desktop, another agent, a CLI. Transport-agnostic: the desktop shell
 * wires this to a stdio pipe, the collab server to a WebSocket.
 *
 * This is only the protocol handler; the tool implementations come from the
 * injected provider (which on the app side dispatches into the DocsBridge).
 */

import type { DocOpsResult, DocOpsTool } from '../agent/coreTypes';
import type { JsonRpcTransport } from './jsonrpc';

const PROTOCOL_VERSION = '2025-06-18';

export interface McpToolProvider {
  listTools(): Promise<DocOpsTool[]> | DocOpsTool[];
  callTool(name: string, args: Record<string, unknown>): Promise<DocOpsResult>;
}

export interface McpServerOptions {
  serverName?: string;
  version?: string;
}

export class McpServer {
  constructor(
    private readonly transport: JsonRpcTransport,
    private readonly provider: McpToolProvider,
    private readonly options: McpServerOptions = {},
  ) {
    transport.onMessage((raw) => void this.handle(raw));
  }

  private reply(id: number, result: unknown): void {
    this.transport.send(JSON.stringify({ jsonrpc: '2.0', id, result }));
  }

  private replyError(id: number, code: number, message: string): void {
    this.transport.send(JSON.stringify({ jsonrpc: '2.0', id, error: { code, message } }));
  }

  private async handle(raw: string): Promise<void> {
    let msg: { id?: number; method?: string; params?: Record<string, unknown> };
    try {
      msg = JSON.parse(raw);
    } catch {
      return;
    }
    const { id, method, params } = msg;

    // Notifications (no id) — nothing to answer.
    if (id === undefined) return;
    if (typeof method !== 'string') {
      this.replyError(id, -32600, 'Invalid request');
      return;
    }

    switch (method) {
      case 'initialize':
        this.reply(id, {
          protocolVersion: PROTOCOL_VERSION,
          capabilities: { tools: {} },
          serverInfo: {
            name: this.options.serverName ?? 'casual-docops',
            version: this.options.version ?? '0.1.0',
          },
        });
        return;

      case 'tools/list': {
        const tools = await this.provider.listTools();
        this.reply(id, {
          tools: tools.map((t) => ({
            name: t.name,
            description: t.description,
            inputSchema: t.input_schema,
          })),
        });
        return;
      }

      case 'tools/call': {
        const name = params?.name;
        const args = (params?.arguments as Record<string, unknown>) ?? {};
        if (typeof name !== 'string') {
          this.replyError(id, -32602, 'tools/call requires a string "name"');
          return;
        }
        const result = await this.provider.callTool(name, args);
        // MCP wraps tool output as content blocks; carry the DocOps result as
        // JSON text and flag errors via isError.
        this.reply(id, {
          content: [{ type: 'text', text: JSON.stringify(result) }],
          isError: !result.ok,
        });
        return;
      }

      default:
        this.replyError(id, -32601, `Method not found: ${method}`);
    }
  }
}
