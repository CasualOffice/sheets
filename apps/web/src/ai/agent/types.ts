/*
 * Copyright (c) 2026 Casual Office. All rights reserved.
 */

/**
 * Types for the agentic DocOps layer — a plan → execute → reflect loop over
 * the ToolRegistry. Deliberately transport-agnostic: the LLM call is injected
 * as `LlmFn`, so the same orchestrator runs against the Anthropic API, the
 * collab server, or the desktop native model.
 */

import type { DocOpsTool } from './coreTypes';

// ── LLM message shapes (Anthropic-compatible) ───────────────────────────────

export type LlmContentBlock =
  | { type: 'text'; text: string }
  | { type: 'tool_use'; id: string; name: string; input: Record<string, unknown> }
  | { type: 'tool_result'; tool_use_id: string; content: string };

export interface LlmMessage {
  role: 'user' | 'assistant';
  content: string | LlmContentBlock[];
}

export interface LlmResponse {
  content: LlmContentBlock[];
  stop_reason: 'end_turn' | 'tool_use' | 'max_tokens' | 'stop_sequence';
}

/** Injected LLM caller. One round-trip; the orchestrator drives the loop. */
export type LlmFn = (args: {
  system: string;
  messages: LlmMessage[];
  tools?: DocOpsTool[];
  /** Streamed text tokens, when the transport supports it. */
  onText?: (token: string) => void;
  signal?: AbortSignal;
}) => Promise<LlmResponse>;

// ── Plan / task model ───────────────────────────────────────────────────────

export type TaskStatus = 'pending' | 'running' | 'done' | 'failed' | 'skipped';

export interface AgentTask {
  id: string;
  /** One-line imperative description shown to the user. */
  title: string;
  status: TaskStatus;
  /** Block ids the task changed, for the diff/accept UI. */
  changedBlockIds?: string[];
  /** Populated when status is 'failed'. */
  error?: string;
}

// ── Events (drive the panel UX) ─────────────────────────────────────────────

export type AgentEvent =
  | { type: 'plan'; tasks: AgentTask[] }
  | { type: 'task-start'; taskId: string }
  | { type: 'task-tool'; taskId: string; tool: string; status: 'running' | 'done' | 'error' }
  | { type: 'task-text'; taskId: string; text: string }
  | { type: 'task-end'; taskId: string; status: TaskStatus; changedBlockIds?: string[] }
  | { type: 'reflect'; goalMet: boolean; note: string; addedTasks: AgentTask[] }
  | { type: 'done'; goalMet: boolean; summary: string }
  | { type: 'error'; message: string };

export interface AgentResult {
  goalMet: boolean;
  tasks: AgentTask[];
  changedBlockIds: string[];
  summary: string;
}

export interface AgentOptions {
  /** Hard cap on planned+reflected tasks. Default 8. */
  maxTasks?: number;
  /** Tool-call rounds per task before giving up. Default 6. */
  maxRoundsPerTask?: number;
  /** Reflection passes that may add corrective tasks. Default 1. */
  maxReflections?: number;
  signal?: AbortSignal;
  onEvent?: (event: AgentEvent) => void;
  /**
   * A short snapshot of the workbook (e.g. sheet names + data extents) given to
   * the PLANNER so it decomposes the goal against real structure instead of
   * guessing. Keep it small (well under the context budget).
   */
  planningContext?: string;
}
