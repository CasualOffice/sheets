/**
 * Macros (Phase 5, T5.1 recorder + T5.2 runner) — record the command-bus
 * mutations a user's edits produce, save them as a named macro, and replay them.
 *
 * We capture `sheet.mutation.*` only: those are the deterministic document state
 * changes (cell values, styles, structural edits). Everything else on the bus —
 * formula-calc triggers, doc rich-text-editing, selection/scroll — is transient
 * noise that must not be replayed. Replay re-executes each mutation through the
 * facade; the formula engine recalculates on its own afterwards.
 *
 * The filter + storage are pure (no @univerjs value imports) so they're unit
 * testable; record/run use the FUniver facade (covered by e2e).
 */
import type { FUniver } from '@univerjs/core/facade';

export type MacroStep = { id: string; params: unknown };
export type Macro = { name: string; steps: MacroStep[]; createdAt: number };

/** True for the deterministic state-change mutations worth recording. */
export function isMacroMutation(id: string): boolean {
  return typeof id === 'string' && id.startsWith('sheet.mutation.');
}

/**
 * Start capturing macro-worthy mutations off the command bus. Returns a `stop`
 * that detaches the listener and yields the recorded steps.
 */
export function startRecording(api: FUniver): { stop: () => MacroStep[] } {
  const steps: MacroStep[] = [];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const disp = (api as any).addEvent((api as any).Event.CommandExecuted, (e: any) => {
    if (e?.id && isMacroMutation(e.id)) steps.push({ id: e.id, params: e.params });
  });
  return {
    stop: () => {
      disp?.dispose?.();
      return steps;
    },
  };
}

/** Replay a macro's steps in order. Best-effort: a failed step is skipped. */
export async function runMacro(api: FUniver, steps: MacroStep[]): Promise<number> {
  let applied = 0;
  for (const s of steps) {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      await (api as any).executeCommand(s.id, s.params);
      applied += 1;
    } catch {
      /* skip a step that no longer applies */
    }
  }
  return applied;
}

const STORAGE_KEY = 'casual.macros';

export function listMacros(): Macro[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? (parsed as Macro[]) : [];
  } catch {
    return [];
  }
}

/** Save (or replace by name) a macro. Returns the updated list. */
export function saveMacro(macro: Macro): Macro[] {
  const next = [...listMacros().filter((m) => m.name !== macro.name), macro];
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
  } catch {
    /* private mode — macro stays in memory only for this session */
  }
  return next;
}

export function deleteMacro(name: string): Macro[] {
  const next = listMacros().filter((m) => m.name !== name);
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
  } catch {
    /* no-op */
  }
  return next;
}

/** Default name for a freshly recorded macro (Macro 1, Macro 2, …). */
export function nextMacroName(existing: Macro[] = listMacros()): string {
  const used = new Set(existing.map((m) => m.name));
  for (let i = 1; ; i += 1) {
    const name = `Macro ${i}`;
    if (!used.has(name)) return name;
  }
}
