/*
 * Copyright (c) 2026 Casual Office. All rights reserved.
 */

/**
 * MCP client exposed as a ToolSource, so external MCP servers (web search,
 * citations, a second document, …) plug straight into the agent's ToolRegistry
 * and become indistinguishable from the built-in DocOps tools.
 *
 * Implements the client half of the Model Context Protocol: initialize
 * handshake → tools/list → tools/call, mapping MCP shapes to our DocOps types.
 */

import type { DocOpsResult, DocOpsTool } from '../agent/coreTypes';
import type { ToolSource } from '../agent/registry';
import { RpcConnection, type JsonRpcTransport } from './jsonrpc';

const PROTOCOL_VERSION = '2025-06-18';

interface McpToolDef {
  name: string;
  description?: string;
  inputSchema?: { type?: string; properties?: Record<string, unknown>; required?: string[] };
}

interface McpCallResult {
  content?: Array<{ type: string; text?: string }>;
  isError?: boolean;
}

export interface McpClientOptions {
  /** Source id in the registry, e.g. 'mcp:search'. */
  id: string;
  clientName?: string;
}

export class McpClient implements ToolSource {
  readonly id: string;
  private readonly rpc: RpcConnection;
  private readonly clientName: string;
  private initialized: Promise<void> | null = null;

  constructor(transport: JsonRpcTransport, options: McpClientOptions) {
    this.id = options.id;
    this.clientName = options.clientName ?? 'casual-docops';
    this.rpc = new RpcConnection(transport);
  }

  /** MCP requires an initialize handshake before any other request. */
  private ensureInitialized(): Promise<void> {
    if (!this.initialized) {
      this.initialized = (async () => {
        await this.rpc.request('initialize', {
          protocolVersion: PROTOCOL_VERSION,
          capabilities: { tools: {} },
          clientInfo: { name: this.clientName, version: '0.1.0' },
        });
        this.rpc.notifyServer('notifications/initialized');
      })();
    }
    return this.initialized;
  }

  async listTools(): Promise<DocOpsTool[]> {
    await this.ensureInitialized();
    const res = await this.rpc.request<{ tools?: McpToolDef[] }>('tools/list');
    return (res.tools ?? []).map((t) => ({
      name: t.name,
      description: t.description ?? t.name,
      input_schema: {
        type: 'object',
        properties: t.inputSchema?.properties ?? {},
        required: t.inputSchema?.required,
      },
    }));
  }

  async callTool(name: string, args: Record<string, unknown>): Promise<DocOpsResult> {
    await this.ensureInitialized();
    try {
      const res = await this.rpc.request<McpCallResult>('tools/call', { name, arguments: args });
      const text = (res.content ?? [])
        .map((c) => (c.type === 'text' ? (c.text ?? '') : ''))
        .join('')
        .trim();
      if (res.isError) {
        return {
          ok: false,
          code: 'UNSUPPORTED',
          message: text || 'MCP tool error',
          retryable: false,
        };
      }
      // External MCP tools don't touch our document model, so there are no
      // changedBlockIds — the text becomes context the model reasons over.
      return { ok: true, data: res.content, diffSummary: text };
    } catch (err) {
      return {
        ok: false,
        code: 'UNSUPPORTED',
        message: err instanceof Error ? err.message : String(err),
        retryable: true,
      };
    }
  }

  close(): void {
    this.rpc.close();
  }
}
