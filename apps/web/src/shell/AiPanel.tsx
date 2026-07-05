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
 * AiPanel — AI assistant task pane for Casual Sheets.
 *
 * Mirrors the DocOpsPanel pattern from the document editor:
 *  - DirectTransport (user API key, browser fetch to Anthropic)
 *  - CollabTransport (server-side LLM loop via /api/ai WS)
 *
 * SheetsBridge translates tool calls into FUniver facade operations.
 * The panel drives the tool loop when drivesLoop=false; the collab
 * server drives it when drivesLoop=true.
 */

import { useCallback, useEffect, useMemo, useRef, useState, type CSSProperties } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { useUniverAPI } from '../use-univer';
import { useUI } from '../use-ui';
import { Icon } from './Icon';
import { SheetsBridge } from '../ai/bridge';
import { SHEETS_CATALOG } from '../ai/catalog';
import { createSheetsTransport, type LlmCallPayload } from '../ai/transport';
import { runAgent, type AgentEvent, type AgentTask, type ToolSource } from '../ai/agent';
import type { McpClient } from '../ai/mcp';
import { createAgentRegistry, createMcpClient, transportLlm } from '../ai/agentRuntime';

interface McpServerState {
  id: string;
  url: string;
  status: 'connecting' | 'connected' | 'error';
  toolCount: number;
  source: McpClient | null;
  error?: string;
}

// ── LLM wire types ─────────────────────────────────────────────────────────

type LlmContentBlock =
  | { type: 'text'; text: string }
  | { type: 'tool_use'; id: string; name: string; input: Record<string, unknown> }
  | { type: 'tool_result'; tool_use_id: string; content: string };

interface LlmMessage {
  role: 'user' | 'assistant';
  content: LlmContentBlock[] | string;
}

interface LlmResponse {
  content: LlmContentBlock[];
  stop_reason: 'end_turn' | 'tool_use' | 'max_tokens' | 'stop_sequence';
}

// ── Display message types ─────────────────────────────────────────────────

type DisplayMessage =
  | { kind: 'user'; text: string }
  | { kind: 'assistant'; text: string }
  | { kind: 'tool_step'; toolName: string; status: 'running' | 'done' | 'error' }
  | { kind: 'error'; text: string }
  | { kind: 'cap'; rounds: number }
  | { kind: 'plan'; tasks: AgentTask[] };

// ── Constants ─────────────────────────────────────────────────────────────

const API_KEY_STORAGE = 'casual_sheets_ai_key';
const MODEL = 'claude-haiku-4-5-20251001';
const DEFAULT_MAX_TOOL_ROUNDS = 12;

const SYSTEM_PROMPT = `You are an AI assistant embedded in Casual Sheets, a spreadsheet app.

You help users read, analyze, and edit their spreadsheets using a structured tool catalog.

Read tools (never mutate):
  get_workbook_info — list all sheets, their names, and dimensions
  get_selection     — read the currently selected range + values
  get_cell_range    — read values from a specific A1 range (e.g. "A1:D10")
  get_sheet_stats   — count rows, columns, and non-empty cells on the active sheet
  find_in_sheet     — search for text or a value (case-insensitive) in the active sheet

Write tools:
  set_cell_values   — write plain text or numbers to a range (2D array, must match range shape)
  set_formula       — write a formula to a single cell (e.g. "=SUM(A1:A10)")

Guidelines:
- Always read before you write. Call get_workbook_info first on a fresh conversation.
- For set_cell_values, call get_cell_range first to confirm the target range is what the user expects.
- For formulas, the leading "=" is optional — both "=SUM(A1:A10)" and "SUM(A1:A10)" are accepted.
- Range notation: plain A1:C3 targets the active sheet; Sheet2!B2:E5 targets a specific sheet.
- Keep responses short. Users want results, not explanations.
- Never invent data about what's in the spreadsheet — always call a read tool first.`;

const TOOL_LABELS: Record<string, string> = {
  get_workbook_info: 'Reading workbook info…',
  get_selection: 'Reading selection…',
  get_cell_range: 'Reading range…',
  get_sheet_stats: 'Computing stats…',
  find_in_sheet: 'Searching…',
  set_cell_values: 'Writing values…',
  set_formula: 'Setting formula…',
};

// ── Styles ────────────────────────────────────────────────────────────────

const messagesStyle: CSSProperties = {
  flex: 1,
  minHeight: 0,
  overflowY: 'auto',
  padding: '10px 14px',
  display: 'flex',
  flexDirection: 'column',
  gap: 8,
};

const msgUserStyle: CSSProperties = {
  alignSelf: 'flex-end',
  maxWidth: '85%',
  background: 'var(--color-primary, #0e7490)',
  color: '#fff',
  borderRadius: '12px 12px 2px 12px',
  padding: '8px 12px',
  fontSize: 13,
  lineHeight: 1.45,
  wordBreak: 'break-word',
};

const msgAssistantStyle: CSSProperties = {
  alignSelf: 'flex-start',
  maxWidth: '95%',
  background: 'var(--color-surface-sunken, #f8f9fa)',
  color: 'var(--color-text)',
  border: '1px solid var(--color-border-light, #e5e7eb)',
  borderRadius: '2px 12px 12px 12px',
  padding: '8px 12px',
  fontSize: 13,
  lineHeight: 1.55,
  wordBreak: 'break-word',
  whiteSpace: 'pre-wrap',
};

const msgToolStyle: CSSProperties = {
  alignSelf: 'flex-start',
  display: 'flex',
  alignItems: 'center',
  gap: 6,
  fontSize: 11.5,
  color: 'var(--color-text-muted, #6b7280)',
  padding: '2px 0',
};

const msgErrorStyle: CSSProperties = {
  alignSelf: 'flex-start',
  maxWidth: '95%',
  background: 'var(--color-danger-bg, #fef2f2)',
  color: 'var(--color-danger, #c62828)',
  border: '1px solid var(--color-danger-border, #fca5a5)',
  borderRadius: 8,
  padding: '8px 12px',
  fontSize: 13,
  lineHeight: 1.45,
};

const msgCapStyle: CSSProperties = {
  alignSelf: 'center',
  fontSize: 11.5,
  color: 'var(--color-text-muted, #6b7280)',
  padding: '3px 10px',
  borderRadius: 6,
  background: 'var(--color-surface-sunken, #f8f9fa)',
  border: '1px solid var(--color-border-light, #e5e7eb)',
};

const msgPlanStyle: CSSProperties = {
  alignSelf: 'stretch',
  display: 'flex',
  flexDirection: 'column',
  gap: 4,
  padding: '8px 12px',
  borderRadius: 8,
  background: 'var(--color-surface-sunken, #f8f9fa)',
  border: '1px solid var(--color-border-light, #e5e7eb)',
};

const msgPlanTitleStyle: CSSProperties = {
  fontSize: 11,
  fontWeight: 600,
  textTransform: 'uppercase',
  letterSpacing: 0.4,
  color: 'var(--color-text-muted, #6b7280)',
  marginBottom: 2,
};

const msgPlanTaskStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  fontSize: 13,
  color: 'var(--color-text, #1f2937)',
};

const agentToggleRowStyle: CSSProperties = {
  display: 'flex',
  padding: '4px 12px 0',
};

const agentToggleStyle = (active: boolean): CSSProperties => ({
  display: 'inline-flex',
  alignItems: 'center',
  gap: 4,
  fontSize: 11.5,
  fontWeight: 600,
  padding: '2px 8px',
  borderRadius: 999,
  cursor: 'pointer',
  color: active ? '#fff' : 'var(--color-text-muted, #6b7280)',
  background: active ? 'var(--color-primary, #1a73e8)' : 'var(--color-surface-sunken, #f8f9fa)',
  border: `1px solid ${active ? 'var(--color-primary, #1a73e8)' : 'var(--color-border-light, #e5e7eb)'}`,
});

const mcpAddBtnStyle: CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  gap: 4,
  fontSize: 11.5,
  fontWeight: 600,
  padding: '2px 8px',
  borderRadius: 999,
  cursor: 'pointer',
  color: 'var(--color-text-muted, #6b7280)',
  background: 'var(--color-surface-sunken, #f8f9fa)',
  border: '1px solid var(--color-border-light, #e5e7eb)',
  marginLeft: 6,
};

const mcpSectionStyle: CSSProperties = {
  display: 'flex',
  flexWrap: 'wrap',
  gap: 6,
  padding: '4px 12px 0',
};

const mcpChipStyle = (status: 'connecting' | 'connected' | 'error'): CSSProperties => ({
  display: 'inline-flex',
  alignItems: 'center',
  gap: 5,
  fontSize: 11,
  padding: '2px 6px',
  borderRadius: 6,
  color: status === 'error' ? 'var(--color-danger, #c62828)' : 'var(--color-text-muted, #6b7280)',
  background: 'var(--color-surface-sunken, #f8f9fa)',
  border: '1px solid var(--color-border-light, #e5e7eb)',
});

const mcpRemoveStyle: CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  background: 'transparent',
  border: 'none',
  cursor: 'pointer',
  padding: 0,
  color: 'inherit',
  opacity: 0.7,
};

const mcpInputStyle: CSSProperties = {
  flex: 1,
  minWidth: 200,
  fontSize: 12,
  padding: '4px 8px',
  borderRadius: 6,
  border: '1px solid var(--color-border-light, #e5e7eb)',
  background: 'var(--color-surface, #fff)',
  color: 'var(--color-text, #1f2937)',
};

const inputRowStyle: CSSProperties = {
  display: 'flex',
  gap: 8,
  padding: '10px 12px',
  borderTop: '1px solid var(--color-border-light, #e5e7eb)',
  alignItems: 'flex-end',
};

// One-tap prompts so the panel isn't a blank chat box — spreadsheet-relevant
// starting points the model + SHEETS_CATALOG tools can act on.
const QUICK_ACTIONS: ReadonlyArray<{ id: string; label: string; prompt: string }> = [
  { id: 'summarize', label: 'Summarize', prompt: 'Summarize the data in this spreadsheet.' },
  {
    id: 'analyze',
    label: 'Analyze',
    prompt: 'Analyze this data and highlight the key trends, totals, and any outliers.',
  },
  {
    id: 'chart',
    label: 'Add chart',
    prompt: 'Suggest and create a chart that best fits this data.',
  },
  {
    id: 'formula',
    label: 'Formula help',
    prompt: 'Suggest a formula to summarize or compute something useful from this data.',
  },
];

const chipRowStyle: CSSProperties = {
  display: 'flex',
  flexWrap: 'wrap',
  gap: 6,
  padding: '0 12px 8px',
};

const chipStyle: CSSProperties = {
  fontSize: 12,
  lineHeight: 1.2,
  padding: '5px 10px',
  border: '1px solid var(--color-border, #d1d5db)',
  borderRadius: 999,
  background: 'var(--color-surface, #ffffff)',
  color: 'var(--color-text, #1f2328)',
  cursor: 'pointer',
  whiteSpace: 'nowrap',
};

const textareaStyle: CSSProperties = {
  flex: 1,
  minWidth: 0,
  fontSize: 13,
  lineHeight: 1.45,
  padding: '8px 10px',
  border: '1px solid var(--color-border, #d1d5db)',
  borderRadius: 8,
  outline: 'none',
  resize: 'none',
  background: 'var(--color-surface, #ffffff)',
  color: 'var(--color-text)',
  font: 'inherit',
  maxHeight: 120,
  overflowY: 'auto',
};

const sendBtnStyle = (busy: boolean): CSSProperties => ({
  padding: '8px 12px',
  borderRadius: 8,
  border: 'none',
  background: busy ? 'var(--color-border, #d1d5db)' : 'var(--color-primary, #0e7490)',
  color: busy ? 'var(--color-text-muted, #6b7280)' : '#fff',
  cursor: busy ? 'not-allowed' : 'pointer',
  flexShrink: 0,
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  transition: 'background 120ms',
});

const keySetupStyle: CSSProperties = {
  flex: 1,
  display: 'flex',
  flexDirection: 'column',
  gap: 12,
  padding: '20px 16px',
};

const keyInputStyle: CSSProperties = {
  width: '100%',
  padding: '8px 10px',
  fontSize: 13,
  border: '1px solid var(--color-border, #d1d5db)',
  borderRadius: 8,
  outline: 'none',
  background: 'var(--color-surface, #ffffff)',
  color: 'var(--color-text)',
  font: 'inherit',
  boxSizing: 'border-box',
};

const saveBtnStyle: CSSProperties = {
  alignSelf: 'flex-start',
  padding: '7px 16px',
  borderRadius: 8,
  border: 'none',
  background: 'var(--color-primary, #0e7490)',
  color: '#fff',
  fontSize: 13,
  fontWeight: 600,
  cursor: 'pointer',
};

const spinnerStyle: CSSProperties = {
  display: 'inline-block',
  width: 10,
  height: 10,
  border: '2px solid currentColor',
  borderTopColor: 'transparent',
  borderRadius: '50%',
  animation: 'ai-panel-spin 0.7s linear infinite',
};

// ── Component ─────────────────────────────────────────────────────────────

export function AiPanel() {
  const ui = useUI();
  const api = useUniverAPI();

  // Keep a mutable ref so the memoized bridge getter always sees the latest
  // FUniver instance without recreating the bridge on every render.
  const apiLatest = useRef<FUniver | null>(null);
  apiLatest.current = api;

  // Stable transport — created once, never re-created on render.
  const transport = useMemo(() => createSheetsTransport(), []);

  // Bridge is stable for the panel's lifetime; reads FUniver through the ref.
  const bridge = useMemo(() => new SheetsBridge(() => apiLatest.current), []);

  const [apiKey, setApiKey] = useState<string>(() => localStorage.getItem(API_KEY_STORAGE) ?? '');
  const [keyDraft, setKeyDraft] = useState('');
  const [showKeySetup, setShowKeySetup] = useState(
    () => transport.requiresApiKey && !localStorage.getItem(API_KEY_STORAGE),
  );

  const [displayMessages, setDisplayMessages] = useState<DisplayMessage[]>([]);
  const [streamingText, setStreamingText] = useState('');
  const [inputValue, setInputValue] = useState('');
  const [busy, setBusy] = useState(false);
  // Agent mode: plan → execute → reflect. Opt-in; only when the panel drives
  // the loop (Direct/Desktop, not collab where the server owns the loop).
  const [agentMode, setAgentMode] = useState(false);
  const [mcpServers, setMcpServers] = useState<McpServerState[]>([]);
  const [mcpUrlDraft, setMcpUrlDraft] = useState('');
  const [showMcpAdd, setShowMcpAdd] = useState(false);

  const historyRef = useRef<LlmMessage[]>([]);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const abortRef = useRef<AbortController | null>(null);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [displayMessages]);

  const appendDisplay = useCallback((msg: DisplayMessage) => {
    setDisplayMessages((prev) => [...prev, msg]);
  }, []);

  const updateLastToolStep = useCallback((status: 'done' | 'error') => {
    setDisplayMessages((prev) => {
      const copy = [...prev];
      for (let i = copy.length - 1; i >= 0; i--) {
        if (copy[i].kind === 'tool_step') {
          copy[i] = { ...(copy[i] as Extract<DisplayMessage, { kind: 'tool_step' }>), status };
          break;
        }
      }
      return copy;
    });
  }, []);

  const updatePlan = useCallback((mutate: (tasks: AgentTask[]) => AgentTask[]) => {
    setDisplayMessages((prev) => {
      const copy = [...prev];
      for (let i = copy.length - 1; i >= 0; i--) {
        if (copy[i].kind === 'plan') {
          const msg = copy[i] as Extract<DisplayMessage, { kind: 'plan' }>;
          copy[i] = { ...msg, tasks: mutate(msg.tasks) };
          break;
        }
      }
      return copy;
    });
  }, []);

  const handleAgentEvent = useCallback(
    (ev: AgentEvent) => {
      switch (ev.type) {
        case 'plan':
          appendDisplay({ kind: 'plan', tasks: ev.tasks });
          break;
        case 'task-start':
          updatePlan((tasks) =>
            tasks.map((t) => (t.id === ev.taskId ? { ...t, status: 'running' } : t)),
          );
          break;
        case 'task-tool':
          if (ev.status === 'running')
            appendDisplay({ kind: 'tool_step', toolName: ev.tool, status: 'running' });
          else updateLastToolStep(ev.status);
          break;
        case 'task-end':
          updatePlan((tasks) =>
            tasks.map((t) => (t.id === ev.taskId ? { ...t, status: ev.status } : t)),
          );
          break;
        case 'reflect':
          if (ev.note) appendDisplay({ kind: 'assistant', text: ev.note });
          if (ev.addedTasks.length) updatePlan((tasks) => [...tasks, ...ev.addedTasks]);
          break;
        case 'error':
          appendDisplay({ kind: 'error', text: ev.message });
          break;
      }
    },
    [appendDisplay, updateLastToolStep, updatePlan],
  );

  const connectMcp = useCallback(
    async (rawUrl: string) => {
      const url = rawUrl.trim();
      if (!url) return;
      const id = `mcp:${url}`;
      if (mcpServers.some((s) => s.id === id)) return;
      const client = createMcpClient(url, id);
      setMcpServers((prev) => [
        ...prev,
        { id, url, status: 'connecting', toolCount: 0, source: client },
      ]);
      setMcpUrlDraft('');
      setShowMcpAdd(false);
      try {
        const tools = await client.listTools();
        setMcpServers((prev) =>
          prev.map((s) =>
            s.id === id ? { ...s, status: 'connected', toolCount: tools.length } : s,
          ),
        );
      } catch (err) {
        setMcpServers((prev) =>
          prev.map((s) =>
            s.id === id
              ? { ...s, status: 'error', error: err instanceof Error ? err.message : String(err) }
              : s,
          ),
        );
      }
    },
    [mcpServers],
  );

  const removeMcp = useCallback((id: string) => {
    setMcpServers((prev) => {
      prev.find((s) => s.id === id)?.source?.close();
      return prev.filter((s) => s.id !== id);
    });
  }, []);

  const saveKey = useCallback(() => {
    const trimmed = keyDraft.trim();
    if (!trimmed) return;
    localStorage.setItem(API_KEY_STORAGE, trimmed);
    setApiKey(trimmed);
    setKeyDraft('');
    setShowKeySetup(false);
  }, [keyDraft]);

  const send = useCallback(
    async (override?: string) => {
      const text = (override ?? inputValue).trim();
      if (!text || busy) return;
      if (transport.requiresApiKey && !apiKey) return;

      setInputValue('');
      setBusy(true);
      appendDisplay({ kind: 'user', text });
      historyRef.current = [...historyRef.current, { role: 'user', content: text }];

      const ctrl = new AbortController();
      abortRef.current = ctrl;

      try {
        if (agentMode && !transport.drivesLoop) {
          // ── Agent mode: plan → execute → reflect ─────────────────────────
          const mcpSources = mcpServers
            .filter((s) => s.status === 'connected' && s.source)
            .map((s) => s.source as ToolSource);
          const registry = createAgentRegistry(bridge, mcpSources);
          const llm = transportLlm(transport, { model: MODEL, apiKey: apiKey || undefined });
          const result = await runAgent(
            text,
            { llm, registry },
            { signal: ctrl.signal, onEvent: handleAgentEvent },
          );
          if (result.summary) {
            appendDisplay({ kind: 'assistant', text: result.summary });
            historyRef.current = [
              ...historyRef.current,
              { role: 'assistant', content: result.summary },
            ];
          }
        } else if (transport.drivesLoop) {
          // ── Server-side loop (CollabTransport) ────────────────────────────
          const payload: LlmCallPayload = {
            model: MODEL,
            max_tokens: 2048,
            system: SYSTEM_PROMPT,
            messages: historyRef.current,
            tools: SHEETS_CATALOG,
            apiKey: apiKey || undefined,
            signal: ctrl.signal,
            maxToolRounds: DEFAULT_MAX_TOOL_ROUNDS,
            toolExecutor: async (toolName, args) => {
              appendDisplay({ kind: 'tool_step', toolName, status: 'running' });
              try {
                const result = await bridge.callTool(toolName, args);
                updateLastToolStep('done');
                return result;
              } catch (err) {
                updateLastToolStep('error');
                throw err;
              }
            },
            onText: (t) => {
              if (t.trim()) appendDisplay({ kind: 'assistant', text: t });
            },
          };

          const { data, status, updatedHistory, capHit } = await transport.call(payload);
          if (status !== 200) {
            const errMsg = (data as { error?: { message?: string } })?.error?.message;
            throw new Error(errMsg ?? `AI error ${status}`);
          }
          if (updatedHistory) historyRef.current = updatedHistory as LlmMessage[];
          if (capHit) appendDisplay({ kind: 'cap', rounds: DEFAULT_MAX_TOOL_ROUNDS });
        } else {
          // ── Panel-driven loop (DirectTransport) ───────────────────────────
          let messages = [...historyRef.current];
          let panelCapHit = false;

          for (let round = 0; round < DEFAULT_MAX_TOOL_ROUNDS; round++) {
            if (ctrl.signal.aborted) break;

            let streamedText = '';
            const payload: LlmCallPayload = {
              model: MODEL,
              max_tokens: 2048,
              system: SYSTEM_PROMPT,
              messages,
              tools: SHEETS_CATALOG,
              apiKey: apiKey || undefined,
              signal: ctrl.signal,
              onText: (tok) => {
                if (tok) {
                  streamedText += tok;
                  setStreamingText((prev) => prev + tok);
                }
              },
            };

            const { data, status } = await transport.call(payload);

            if (streamedText.trim()) {
              appendDisplay({ kind: 'assistant', text: streamedText });
            }
            setStreamingText('');

            if (status !== 200) {
              const errMsg = (data as { error?: { message?: string } })?.error?.message;
              throw new Error(errMsg ?? `API error ${status}`);
            }

            const response = data as LlmResponse;
            messages = [...messages, { role: 'assistant', content: response.content }];

            if (!streamedText) {
              for (const block of response.content) {
                if (block.type === 'text' && block.text.trim()) {
                  appendDisplay({ kind: 'assistant', text: block.text });
                }
              }
            }

            if (response.stop_reason !== 'tool_use') break;

            const toolUses = response.content.filter(
              (b): b is Extract<LlmContentBlock, { type: 'tool_use' }> => b.type === 'tool_use',
            );
            const toolResults: LlmContentBlock[] = [];

            for (const tu of toolUses) {
              appendDisplay({ kind: 'tool_step', toolName: tu.name, status: 'running' });
              try {
                const result = await bridge.callTool(tu.name, tu.input);
                updateLastToolStep('done');
                toolResults.push({
                  type: 'tool_result',
                  tool_use_id: tu.id,
                  content: JSON.stringify(result),
                });
              } catch (err) {
                updateLastToolStep('error');
                toolResults.push({
                  type: 'tool_result',
                  tool_use_id: tu.id,
                  content: JSON.stringify({ ok: false, message: String(err) }),
                });
              }
            }

            messages = [...messages, { role: 'user', content: toolResults }];

            if (round === DEFAULT_MAX_TOOL_ROUNDS - 1) panelCapHit = true;
          }

          historyRef.current = messages;
          if (panelCapHit) appendDisplay({ kind: 'cap', rounds: DEFAULT_MAX_TOOL_ROUNDS });
        }
      } catch (err) {
        if ((err as { name?: string }).name === 'AbortError') return;
        const msg = err instanceof Error ? err.message : String(err);
        appendDisplay({ kind: 'error', text: msg });
      } finally {
        setBusy(false);
        abortRef.current = null;
      }
    },
    [
      inputValue,
      busy,
      transport,
      apiKey,
      bridge,
      appendDisplay,
      updateLastToolStep,
      agentMode,
      handleAgentEvent,
      mcpServers,
    ],
  );

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        void send();
      }
    },
    [send],
  );

  const handleClose = useCallback(() => {
    abortRef.current?.abort();
    ui.toggleAiPanel();
  }, [ui]);

  return (
    <>
      <style>{`@keyframes ai-panel-spin { to { transform: rotate(360deg); } }`}</style>
      <aside className="side-panel ai-panel" data-testid="ai-panel">
        <header className="side-panel__header">
          <Icon name="auto_awesome" />
          <h2 className="side-panel__title">AI</h2>
          <button
            type="button"
            className="side-panel__close"
            aria-label="Close AI panel"
            onClick={handleClose}
          >
            <Icon name="close" />
          </button>
        </header>

        <div
          className="side-panel__body"
          style={{ display: 'flex', flexDirection: 'column', minHeight: 0, flex: 1 }}
        >
          {showKeySetup ? (
            <div style={keySetupStyle}>
              <p
                style={{ fontSize: 13, color: 'var(--color-text-muted, #6b7280)', lineHeight: 1.5 }}
              >
                Enter your Anthropic API key to use the AI assistant. The key is stored locally in
                your browser and never sent anywhere except Anthropic.
              </p>
              <input
                type="password"
                placeholder="sk-ant-..."
                value={keyDraft}
                style={keyInputStyle}
                onChange={(e) => setKeyDraft(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && saveKey()}
                aria-label="Anthropic API key"
              />
              <button type="button" style={saveBtnStyle} onClick={saveKey}>
                Save key
              </button>
              {apiKey && (
                <button
                  type="button"
                  style={{
                    ...saveBtnStyle,
                    background: 'transparent',
                    color: 'var(--color-primary, #0e7490)',
                    border: '1px solid var(--color-primary, #0e7490)',
                  }}
                  onClick={() => setShowKeySetup(false)}
                >
                  Keep existing key
                </button>
              )}
            </div>
          ) : (
            <>
              <div style={messagesStyle}>
                {displayMessages.length === 0 && (
                  <div
                    style={{
                      color: 'var(--color-text-muted, #6b7280)',
                      fontSize: 13,
                      textAlign: 'center',
                      marginTop: 24,
                    }}
                  >
                    Ask anything about this spreadsheet, or ask me to make changes.
                  </div>
                )}
                {displayMessages.map((msg, i) => {
                  if (msg.kind === 'user') {
                    return (
                      <div key={i} style={msgUserStyle}>
                        {msg.text}
                      </div>
                    );
                  }
                  if (msg.kind === 'assistant') {
                    return (
                      <div key={i} style={msgAssistantStyle}>
                        {msg.text}
                      </div>
                    );
                  }
                  if (msg.kind === 'tool_step') {
                    return (
                      <div key={i} style={msgToolStyle}>
                        {msg.status === 'running' ? (
                          <span style={spinnerStyle} />
                        ) : msg.status === 'done' ? (
                          <Icon name="check_circle" />
                        ) : (
                          <Icon name="error" />
                        )}
                        <span>{TOOL_LABELS[msg.toolName] ?? msg.toolName}</span>
                      </div>
                    );
                  }
                  if (msg.kind === 'error') {
                    return (
                      <div key={i} style={msgErrorStyle}>
                        {msg.text}
                      </div>
                    );
                  }
                  if (msg.kind === 'cap') {
                    return (
                      <div key={i} style={msgCapStyle}>
                        Stopped after {msg.rounds} tool steps — send another message to continue.
                      </div>
                    );
                  }
                  if (msg.kind === 'plan') {
                    return (
                      <div key={i} style={msgPlanStyle} data-testid="ai-plan">
                        <div style={msgPlanTitleStyle}>Plan</div>
                        {msg.tasks.map((t) => (
                          <div key={t.id} style={msgPlanTaskStyle}>
                            {t.status === 'running' ? (
                              <span style={spinnerStyle} aria-hidden="true" />
                            ) : t.status === 'done' ? (
                              <Icon name="check" />
                            ) : t.status === 'failed' ? (
                              <Icon name="close" />
                            ) : (
                              <Icon name="radio_button_unchecked" />
                            )}
                            <span style={{ opacity: t.status === 'pending' ? 0.6 : 1 }}>
                              {t.title}
                            </span>
                          </div>
                        ))}
                      </div>
                    );
                  }
                  return null;
                })}
                {streamingText && (
                  <div style={{ ...msgAssistantStyle, opacity: 0.85 }}>
                    {streamingText}
                    <span style={spinnerStyle} aria-hidden="true" />
                  </div>
                )}
                <div ref={messagesEndRef} />
              </div>

              {!busy && !(transport.requiresApiKey && !apiKey) && (
                <div style={chipRowStyle}>
                  {QUICK_ACTIONS.map((a) => (
                    <button
                      key={a.id}
                      type="button"
                      style={chipStyle}
                      onClick={() => void send(a.prompt)}
                      data-testid={`ai-quick-${a.id}`}
                    >
                      {a.label}
                    </button>
                  ))}
                </div>
              )}

              {!transport.drivesLoop && (
                <div style={agentToggleRowStyle}>
                  <button
                    type="button"
                    onClick={() => setAgentMode((v) => !v)}
                    style={agentToggleStyle(agentMode)}
                    data-testid="ai-agent-toggle"
                    title={
                      agentMode
                        ? 'Agent mode — plans, executes, and reviews multi-step tasks'
                        : 'Chat mode — single reply'
                    }
                    disabled={busy}
                  >
                    <Icon name="smart_toy" />
                    {agentMode ? 'Agent' : 'Chat'}
                  </button>
                  {agentMode && (
                    <button
                      type="button"
                      onClick={() => setShowMcpAdd((v) => !v)}
                      style={mcpAddBtnStyle}
                      data-testid="ai-mcp-add"
                      title="Connect an external MCP server; its tools join the agent"
                    >
                      <Icon name="hub" />
                      MCP
                    </button>
                  )}
                </div>
              )}
              {agentMode && !transport.drivesLoop && (mcpServers.length > 0 || showMcpAdd) && (
                <div style={mcpSectionStyle} data-testid="ai-mcp-section">
                  {mcpServers.map((s) => (
                    <div key={s.id} style={mcpChipStyle(s.status)}>
                      {s.status === 'connecting' ? (
                        <span style={spinnerStyle} aria-hidden="true" />
                      ) : s.status === 'connected' ? (
                        <Icon name="check_circle" />
                      ) : (
                        <Icon name="error" />
                      )}
                      <span title={s.error ?? s.url}>
                        {s.url.replace(/^https?:\/\//, '')}
                        {s.status === 'connected' ? ` · ${s.toolCount} tools` : ''}
                        {s.status === 'error' ? ' · failed' : ''}
                      </span>
                      <button
                        type="button"
                        onClick={() => removeMcp(s.id)}
                        style={mcpRemoveStyle}
                        aria-label="Remove MCP server"
                      >
                        <Icon name="close" />
                      </button>
                    </div>
                  ))}
                  {showMcpAdd && (
                    <input
                      value={mcpUrlDraft}
                      onChange={(e) => setMcpUrlDraft(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          e.preventDefault();
                          void connectMcp(mcpUrlDraft);
                        }
                      }}
                      placeholder="https://mcp.example.com/rpc  (Enter to connect)"
                      style={mcpInputStyle}
                      data-testid="ai-mcp-input"
                    />
                  )}
                </div>
              )}
              <div style={inputRowStyle}>
                <textarea
                  rows={1}
                  placeholder="Ask about this spreadsheet…"
                  value={inputValue}
                  style={textareaStyle}
                  disabled={busy}
                  onChange={(e) => setInputValue(e.target.value)}
                  onKeyDown={handleKeyDown}
                  aria-label="Message input"
                />
                <button
                  type="button"
                  style={sendBtnStyle(busy)}
                  disabled={busy}
                  onClick={() => void send()}
                  aria-label="Send"
                >
                  {busy ? <span style={spinnerStyle} /> : <Icon name="send" />}
                </button>
              </div>

              {transport.requiresApiKey && (
                <button
                  type="button"
                  style={{
                    fontSize: 11,
                    color: 'var(--color-text-muted, #6b7280)',
                    background: 'none',
                    border: 'none',
                    cursor: 'pointer',
                    padding: '4px 12px 8px',
                    textDecoration: 'underline',
                  }}
                  onClick={() => setShowKeySetup(true)}
                >
                  Change API key
                </button>
              )}
            </>
          )}
        </div>
      </aside>
    </>
  );
}
