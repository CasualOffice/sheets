/*
 * Copyright (c) 2026 Casual Office. All rights reserved.
 */

/**
 * The agentic DocOps loop: plan → execute → reflect.
 *
 *   plan     — decompose the user's goal into an ordered list of concrete
 *              sub-tasks (one LLM call, `submit_plan` tool).
 *   execute  — run each sub-task through a ReAct tool loop over the registry;
 *              every mutation flows back as a DocOps result (tracked change).
 *   reflect  — re-assess whether the goal is met; if not, append corrective
 *              tasks and execute them (bounded by maxReflections).
 *
 * Transport-agnostic (LLM injected as LlmFn) and tool-source-agnostic (tools
 * come from the ToolRegistry, so built-in + MCP tools are indistinguishable).
 */

import type { DocOpsTool } from './coreTypes';
import type { ToolRegistry } from './registry';
import type {
  AgentEvent,
  AgentOptions,
  AgentResult,
  AgentTask,
  LlmContentBlock,
  LlmFn,
  LlmMessage,
} from './types';

// ── Meta-tools the agent uses to talk to itself ─────────────────────────────

const SUBMIT_PLAN: DocOpsTool = {
  name: 'submit_plan',
  description:
    'Submit the ordered list of concrete sub-tasks that accomplish the goal. Each task is one imperative sentence a document editor can carry out (e.g. "Rewrite the introduction to be more formal").',
  input_schema: {
    type: 'object',
    properties: {
      tasks: {
        type: 'array',
        items: { type: 'string' },
        description: '2–6 ordered sub-tasks. Keep them concrete and independent.',
      },
    },
    required: ['tasks'],
  },
};

const SUBMIT_REFLECTION: DocOpsTool = {
  name: 'submit_reflection',
  description:
    'Report whether the goal has now been fully accomplished. If not, list the remaining corrective sub-tasks.',
  input_schema: {
    type: 'object',
    properties: {
      goalMet: { type: 'boolean', description: 'True only if the goal is fully accomplished.' },
      note: { type: 'string', description: 'One sentence explaining the assessment.' },
      remainingTasks: {
        type: 'array',
        items: { type: 'string' },
        description: 'Corrective sub-tasks if goalMet is false; empty otherwise.',
      },
    },
    required: ['goalMet'],
  },
};

const PLANNER_SYSTEM =
  'You are the planner for an AI document assistant. Break the user goal into a short ordered list of concrete, independently-executable editing sub-tasks, then call submit_plan. Do not perform the edits yourself. Prefer the fewest tasks that fully cover the goal.';

const REFLECT_SYSTEM =
  'You are the reviewer for an AI document assistant. Given the original goal and a log of the edits performed, judge whether the goal is fully accomplished. Call submit_reflection with your verdict and any remaining corrective sub-tasks.';

const executorSystem = (goal: string, task: string): string =>
  `You are the executor for ONE step of a document task.
Overall goal: ${goal}
Your current step: ${task}

You accomplish the step ONLY by calling tools — you cannot edit the document any other way. NEVER claim an edit is done without calling the tool that performs it.

Procedure:
1. First call a read tool (get_outline, get_selection, or find_text) to locate the target.
2. Then call the write tool that makes the change (e.g. rewrite_selection, insert_paragraph_after, set_paragraph_style).
Emit each call as:
<tool_call>
{"name": "<tool>", "arguments": { ... }}
</tool_call>
Do NOT describe the edit in prose instead of calling the tool. Only after a write tool returns success may you reply with a one-line confirmation and stop. Do not start other steps.`;

// ── Small helpers ───────────────────────────────────────────────────────────

function textOf(content: LlmContentBlock[]): string {
  return content
    .filter((b): b is Extract<LlmContentBlock, { type: 'text' }> => b.type === 'text')
    .map((b) => b.text)
    .join('')
    .trim();
}

function toolUses(content: LlmContentBlock[]): Extract<LlmContentBlock, { type: 'tool_use' }>[] {
  return content.filter(
    (b): b is Extract<LlmContentBlock, { type: 'tool_use' }> => b.type === 'tool_use',
  );
}

/** Parse a plan from a tool_use, else fall back to numbered/bulleted text. */
function parseTasks(content: LlmContentBlock[], toolName: string): string[] {
  const call = toolUses(content).find((t) => t.name === toolName);
  const raw = (call?.input?.tasks ?? call?.input?.remainingTasks) as unknown;
  if (Array.isArray(raw)) {
    return raw.map((t) => String(t).trim()).filter(Boolean);
  }
  // Fallback: the model wrote a list as prose.
  return textOf(content)
    .split('\n')
    .map((l) => l.replace(/^\s*(?:[-*]|\d+[.)])\s*/, '').trim())
    .filter((l) => l.length > 0);
}

let taskSeq = 0;
function makeTask(title: string): AgentTask {
  taskSeq += 1;
  return { id: `t${taskSeq}`, title, status: 'pending' };
}

// ── The orchestrator ────────────────────────────────────────────────────────

export async function runAgent(
  goal: string,
  deps: { llm: LlmFn; registry: ToolRegistry },
  options: AgentOptions = {},
): Promise<AgentResult> {
  const { llm, registry } = deps;
  const maxTasks = options.maxTasks ?? 8;
  const maxRoundsPerTask = options.maxRoundsPerTask ?? 6;
  const maxReflections = options.maxReflections ?? 1;
  const signal = options.signal;
  const emit = (e: AgentEvent): void => options.onEvent?.(e);

  const changed = new Set<string>();
  const executionLog: string[] = [];

  const aborted = (): boolean => !!signal?.aborted;

  try {
    // ── 1. PLAN ──────────────────────────────────────────────────────────
    // Ground the planner with a cheap workbook snapshot when supplied, so the
    // plan reflects real structure (a "summarize" goal isn't turned into edits).
    const planUserMsg = options.planningContext
      ? `Workbook snapshot:\n${options.planningContext}\n\nGoal: ${goal}`
      : goal;
    const planResp = await llm({
      system: PLANNER_SYSTEM,
      messages: [{ role: 'user', content: planUserMsg }],
      tools: [SUBMIT_PLAN],
      signal,
    });
    let titles = parseTasks(planResp.content, 'submit_plan').slice(0, maxTasks);
    if (titles.length === 0) titles = [goal]; // degenerate: treat the goal as one task
    const tasks: AgentTask[] = titles.map(makeTask);
    emit({ type: 'plan', tasks: tasks.map((t) => ({ ...t })) });

    // ── 2. EXECUTE (with bounded reflection) ─────────────────────────────
    let reflections = 0;
    // Index of the first not-yet-executed task; reflection appends more.
    let cursor = 0;
    while (cursor < tasks.length && !aborted()) {
      const task = tasks[cursor];
      cursor += 1;
      task.status = 'running';
      emit({ type: 'task-start', taskId: task.id });

      const outcome = await executeTask(task, goal, llm, registry, {
        maxRounds: maxRoundsPerTask,
        signal,
        emit,
      });
      outcome.changed.forEach((id) => changed.add(id));
      task.changedBlockIds = outcome.changed;
      task.status = outcome.failed ? 'failed' : 'done';
      if (outcome.failed) task.error = outcome.error;
      executionLog.push(`[${task.status}] ${task.title}`);
      emit({
        type: 'task-end',
        taskId: task.id,
        status: task.status,
        changedBlockIds: outcome.changed,
      });

      // When the plan is exhausted, reflect once (or up to maxReflections) and
      // let it append corrective tasks before we declare completion.
      if (cursor >= tasks.length && reflections < maxReflections && !aborted()) {
        reflections += 1;
        const verdict = await reflect(goal, executionLog, llm, signal);
        const added = verdict.remaining.slice(0, maxTasks - tasks.length).map(makeTask);
        tasks.push(...added);
        emit({
          type: 'reflect',
          goalMet: verdict.goalMet,
          note: verdict.note,
          addedTasks: added.map((t) => ({ ...t })),
        });
        if (verdict.goalMet || added.length === 0) break;
      }
    }

    const goalMet = tasks.every((t) => t.status === 'done' || t.status === 'skipped');
    const summary = `${tasks.filter((t) => t.status === 'done').length}/${tasks.length} steps completed.`;
    emit({ type: 'done', goalMet, summary });
    return { goalMet, tasks, changedBlockIds: [...changed], summary };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    emit({ type: 'error', message });
    return {
      goalMet: false,
      tasks: [],
      changedBlockIds: [...changed],
      summary: message,
    };
  }
}

interface TaskOutcome {
  changed: string[];
  failed: boolean;
  error?: string;
}

/** Run one sub-task through a ReAct tool loop. */
async function executeTask(
  task: AgentTask,
  goal: string,
  llm: LlmFn,
  registry: ToolRegistry,
  opts: { maxRounds: number; signal?: AbortSignal; emit: (e: AgentEvent) => void },
): Promise<TaskOutcome> {
  const tools = await registry.tools();
  const messages: LlmMessage[] = [{ role: 'user', content: task.title }];
  const changed: string[] = [];
  let lastError: string | undefined;
  let attemptedTools = 0;
  let failedTools = 0;

  for (let round = 0; round < opts.maxRounds; round++) {
    if (opts.signal?.aborted) return { changed, failed: true, error: 'Cancelled.' };
    const resp = await llm({
      system: executorSystem(goal, task.title),
      messages,
      tools,
      onText: (t) => opts.emit({ type: 'task-text', taskId: task.id, text: t }),
      signal: opts.signal,
    });

    const calls = toolUses(resp.content);
    // Keep only tool_use blocks when continuing so history stays balanced.
    messages.push({
      role: 'assistant',
      content: resp.stop_reason === 'tool_use' ? resp.content : textOf(resp.content) || ' ',
    });
    if (resp.stop_reason !== 'tool_use' || calls.length === 0) break;

    const results: LlmContentBlock[] = [];
    for (const call of calls) {
      opts.emit({ type: 'task-tool', taskId: task.id, tool: call.name, status: 'running' });
      const result = await registry.call(call.name, call.input);
      attemptedTools += 1;
      if (result.ok) {
        if (result.changedBlockIds) changed.push(...result.changedBlockIds);
      } else {
        failedTools += 1;
        lastError = result.message;
      }
      opts.emit({
        type: 'task-tool',
        taskId: task.id,
        tool: call.name,
        status: result.ok ? 'done' : 'error',
      });
      results.push({
        type: 'tool_result',
        tool_use_id: call.id,
        content: JSON.stringify(result),
      });
    }
    messages.push({ role: 'user', content: results });
  }

  // A task is a failure only if it attempted tools and every one errored.
  const failed = attemptedTools > 0 && failedTools === attemptedTools;
  return { changed, failed, error: failed ? lastError : undefined };
}

/** Reflection pass: did we meet the goal, and what's left? */
async function reflect(
  goal: string,
  log: string[],
  llm: LlmFn,
  signal?: AbortSignal,
): Promise<{ goalMet: boolean; note: string; remaining: string[] }> {
  const resp = await llm({
    system: REFLECT_SYSTEM,
    messages: [
      {
        role: 'user',
        content: `Goal: ${goal}\n\nEdits performed:\n${log.join('\n') || '(none)'}\n\nAssess completion and call submit_reflection.`,
      },
    ],
    tools: [SUBMIT_REFLECTION],
    signal,
  });
  const call = toolUses(resp.content).find((t) => t.name === 'submit_reflection');
  const goalMet = call?.input?.goalMet === true;
  const note = typeof call?.input?.note === 'string' ? (call.input.note as string) : '';
  const remaining = goalMet ? [] : parseTasks(resp.content, 'submit_reflection');
  return { goalMet, note, remaining };
}
