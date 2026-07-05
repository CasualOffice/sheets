/*
 * Copyright (c) 2026 Casual Office. All rights reserved.
 */

/**
 * Tool registry — the single seam through which the agent reaches every tool,
 * regardless of origin. This is what makes the system MCP-pluggable: the
 * built-in DocOps catalog is one `ToolSource`; an MCP client connected to an
 * external server (web search, citations, another document) is another. The
 * agent orchestrator never knows or cares which source owns a tool.
 */

import type { DocOpsResult, DocOpsTool } from './coreTypes';

/**
 * A provider of callable tools. The built-in bridge implements this over the
 * in-process DocOps catalog; an MCP client implements it over a remote server.
 */
export interface ToolSource {
  /** Stable id for diagnostics and prefixing (e.g. 'docops', 'mcp:search'). */
  readonly id: string;
  /** The tools this source exposes, in Anthropic tool-definition shape. */
  listTools(): Promise<DocOpsTool[]> | DocOpsTool[];
  /** Execute one of this source's tools. */
  callTool(name: string, args: Record<string, unknown>): Promise<DocOpsResult>;
}

/**
 * Merges tools from every registered source into one catalog and routes each
 * call to the owning source. Name collisions are resolved first-registered-wins
 * and reported via `collisions` so the caller can warn.
 */
export class ToolRegistry {
  private readonly sources: ToolSource[] = [];
  /** Cache of tool name → owning source, rebuilt on each `tools()`. */
  private owner = new Map<string, ToolSource>();
  readonly collisions: string[] = [];

  register(source: ToolSource): void {
    if (this.sources.some((s) => s.id === source.id)) {
      throw new Error(`ToolSource '${source.id}' is already registered`);
    }
    this.sources.push(source);
  }

  /** The merged tool catalog. Rebuilds the routing table as a side effect. */
  async tools(): Promise<DocOpsTool[]> {
    this.owner.clear();
    this.collisions.length = 0;
    const merged: DocOpsTool[] = [];
    for (const source of this.sources) {
      const tools = await source.listTools();
      for (const tool of tools) {
        if (this.owner.has(tool.name)) {
          this.collisions.push(tool.name);
          continue; // first-registered wins
        }
        this.owner.set(tool.name, source);
        merged.push(tool);
      }
    }
    return merged;
  }

  /**
   * Route a call to the source that owns `name`. Refreshes the routing table
   * first if the tool is unknown (a source may have added tools since the last
   * `tools()`), then returns a structured error rather than throwing so the
   * agent loop can feed it back to the model.
   */
  async call(name: string, args: Record<string, unknown>): Promise<DocOpsResult> {
    if (!this.owner.has(name)) {
      await this.tools();
    }
    const source = this.owner.get(name);
    if (!source) {
      return {
        ok: false,
        code: 'UNSUPPORTED',
        message: `No registered tool named '${name}'.`,
        retryable: false,
      };
    }
    try {
      return await source.callTool(name, args);
    } catch (err) {
      return {
        ok: false,
        code: 'UNSUPPORTED',
        message: err instanceof Error ? err.message : String(err),
        retryable: false,
      };
    }
  }
}
