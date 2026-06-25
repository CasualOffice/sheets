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
export type Macro = {
  name: string;
  steps: MacroStep[];
  createdAt: number;
  /** Optional Ctrl+Shift+<letter> binding (single uppercase A–Z). */
  shortcut?: string;
};

/**
 * Ctrl+Shift+<letter> combos the app already owns (insert-table-as, clear
 * formatting, command palette). Macros can't claim these, else the binding
 * would fight a built-in shortcut. Keep in sync with MenuBar's keydown.
 */
export const RESERVED_MACRO_LETTERS: ReadonlySet<string> = new Set(['L', 'D', 'P']);

/** Letters a macro may bind to: A–Z minus reserved minus those another macro holds. */
export function availableMacroLetters(name: string, macros: Macro[] = listMacros()): string[] {
  const taken = new Set(
    macros.filter((m) => m.name !== name && m.shortcut).map((m) => m.shortcut as string),
  );
  const letters: string[] = [];
  for (let c = 65; c <= 90; c += 1) {
    const ch = String.fromCharCode(c);
    if (!RESERVED_MACRO_LETTERS.has(ch) && !taken.has(ch)) letters.push(ch);
  }
  return letters;
}

/**
 * Assign (or clear, with `null`) a macro's Ctrl+Shift+<letter> binding. A
 * shortcut is unique: any other macro holding the same letter loses it.
 * Reserved letters are rejected (binding left unchanged). Returns the list.
 */
export function setMacroShortcut(name: string, letter: string | null): Macro[] {
  const up = letter ? letter.toUpperCase() : null;
  if (up && (up.length !== 1 || up < 'A' || up > 'Z' || RESERVED_MACRO_LETTERS.has(up))) {
    return listMacros();
  }
  const next = listMacros().map((m) => {
    if (m.name === name) return { ...m, shortcut: up ?? undefined };
    if (up && m.shortcut === up) return { ...m, shortcut: undefined };
    return m;
  });
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
  } catch {
    /* private mode — change stays in memory only for this session */
  }
  return next;
}

/** The macro bound to Ctrl+Shift+<letter>, if any. */
export function findMacroByShortcut(letter: string): Macro | undefined {
  const up = letter.toUpperCase();
  return listMacros().find((m) => m.shortcut === up);
}

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

/**
 * Rename a macro in place (preserving list order, steps, shortcut, createdAt).
 * No-op — returns the list unchanged — when the new name is blank, unchanged,
 * collides with another macro, or the old name isn't found. The name is the
 * storage key, so the rename is a keyed map, not a delete + re-add.
 */
export function renameMacro(oldName: string, newName: string): Macro[] {
  const trimmed = newName.trim();
  const macros = listMacros();
  if (!trimmed || trimmed === oldName) return macros;
  if (!macros.some((m) => m.name === oldName)) return macros;
  if (macros.some((m) => m.name === trimmed)) return macros; // name collision
  const next = macros.map((m) => (m.name === oldName ? { ...m, name: trimmed } : m));
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
  } catch {
    /* private mode — change stays in memory only for this session */
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
