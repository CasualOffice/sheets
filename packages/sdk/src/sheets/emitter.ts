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
 * Minimal typed event emitter backing the SDK's canonical `.on()/.off()` surface
 * (doc 38 §3). Kept in its own Univer-free module so it can be unit-tested in the
 * plain node test runner (importing `./api` would pull the Univer runtime).
 */

/** Base constraint for an event map — a record of handler signatures. */
export type EventFn = (...args: never[]) => void;

export interface Emitter<E extends Record<keyof E, EventFn>> {
  /** Subscribe. Returns an unsubscribe function. */
  on<K extends keyof E>(name: K, handler: E[K]): () => void;
  /** Remove a previously-registered handler. */
  off<K extends keyof E>(name: K, handler: E[K]): void;
  /** Dispatch to all current subscribers of `name`. */
  emit<K extends keyof E>(name: K, ...args: Parameters<E[K]>): void;
  /** Live subscriber count for `name`. */
  listenerCount<K extends keyof E>(name: K): number;
}

/**
 * Build a typed emitter. `emit` snapshots the handler set so a handler that
 * unsubscribes mid-dispatch doesn't skip its peers, and isolates handler throws
 * so a misbehaving host callback can never break the seam `emit` is called from
 * (e.g. Univer's mutation / selection callbacks).
 */
export function createEmitter<E extends Record<keyof E, EventFn>>(): Emitter<E> {
  const listeners = new Map<keyof E, Set<EventFn>>();
  return {
    on<K extends keyof E>(name: K, handler: E[K]): () => void {
      let set = listeners.get(name);
      if (!set) {
        set = new Set();
        listeners.set(name, set);
      }
      set.add(handler as EventFn);
      return () => {
        listeners.get(name)?.delete(handler as EventFn);
      };
    },
    off<K extends keyof E>(name: K, handler: E[K]): void {
      listeners.get(name)?.delete(handler as EventFn);
    },
    emit<K extends keyof E>(name: K, ...args: Parameters<E[K]>): void {
      const set = listeners.get(name);
      if (!set) return;
      for (const handler of [...set]) {
        try {
          (handler as (...a: Parameters<E[K]>) => void)(...args);
        } catch (err) {
          if (typeof console !== 'undefined') {
            console.error(`[CasualSheets] "${String(name)}" event handler threw`, err);
          }
        }
      }
    },
    listenerCount<K extends keyof E>(name: K): number {
      return listeners.get(name)?.size ?? 0;
    },
  };
}
