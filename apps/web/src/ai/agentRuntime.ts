/*
 * Copyright (c) 2026 Casual Office. All rights reserved.
 */

/**
 * Wires the transport-agnostic agent + MCP core (ai/agent, ai/mcp — ported from
 * @casualoffice/docops) to the Sheets runtime: SheetsBridge becomes a
 * ToolSource, the SheetsTransport becomes the injected LlmFn, and external MCP
 * servers can be registered into the same ToolRegistry.
 */

import { SHEETS_CATALOG } from './catalog';
import type { SheetsBridge } from './bridge';
import type { SheetsTransport } from './transport';
import { ToolRegistry, type LlmFn, type LlmResponse, type ToolSource } from './agent';
import { HttpMcpTransport, McpClient } from './mcp';
import type { DocOpsTool } from './agent/coreTypes';

/** Adapt the in-process SheetsBridge to a ToolSource. */
export function bridgeToolSource(bridge: SheetsBridge): ToolSource {
  return {
    id: 'sheets',
    listTools: () => SHEETS_CATALOG as DocOpsTool[],
    callTool: (name, args) => bridge.callTool(name, args),
  };
}

/** Built-in Sheets tools first (win collisions), then external MCP sources. */
export function createAgentRegistry(bridge: SheetsBridge, extra: ToolSource[] = []): ToolRegistry {
  const registry = new ToolRegistry();
  registry.register(bridgeToolSource(bridge));
  for (const source of extra) registry.register(source);
  return registry;
}

/** Connect to an external MCP server over Streamable HTTP as a ToolSource. */
export function createMcpClient(
  url: string,
  id: string,
  headers?: Record<string, string>,
  proxyUrl?: string,
): McpClient {
  return new McpClient(new HttpMcpTransport(url, { headers, proxyUrl }), { id });
}

/**
 * Adapt a single-round SheetsTransport (Direct/Desktop, drivesLoop=false) to the
 * agent's LlmFn. Throws on non-200 so the agent surfaces the real error.
 */
export function transportLlm(
  transport: SheetsTransport,
  opts: { model: string; apiKey?: string; maxTokens?: number },
): LlmFn {
  return async ({ system, messages, tools, onText, signal }) => {
    const { data, status } = await transport.call({
      model: opts.model,
      max_tokens: opts.maxTokens ?? 2048,
      system,
      messages,
      tools: tools ?? [],
      apiKey: opts.apiKey || undefined,
      signal,
      onText,
    });
    if (status !== 200) {
      throw new Error(friendlyLlmError(status, data));
    }
    return data as LlmResponse;
  };
}

/**
 * Turn an HTTP status + Anthropic-shaped error body into an actionable message,
 * so users see "your key was rejected" instead of a bare "API error 401".
 */
export function friendlyLlmError(status: number, data: unknown): string {
  const raw = (data as { error?: { message?: string } })?.error?.message;
  if (status === 401 || status === 403)
    return raw || 'Your API key was rejected — check it in AI settings.';
  if (status === 429) return raw || 'Rate limited — wait a moment and try again.';
  if (status >= 500) return raw || `The AI service is unavailable (error ${status}).`;
  return raw || `AI request failed (error ${status}).`;
}
