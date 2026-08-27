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

import { clonePlainCommandData } from './command-context';

type ExecutionOptions = Record<PropertyKey, unknown>;
const PUBLIC_EXECUTION_METHODS = new Set<PropertyKey>(['executeCommand', 'syncExecuteCommand']);
const PUBLIC_COMMAND_LISTENER_METHODS = new Set<PropertyKey>([
  'onBeforeCommandExecute',
  'onCommandExecuted',
]);
const PUBLIC_COMMAND_EVENTS = new Set(['BeforeCommandExecute', 'CommandExecuted']);

type PublicCallback = (...args: unknown[]) => unknown;

function isPromiseLike(value: unknown): value is PromiseLike<unknown> {
  return (
    value !== null &&
    (typeof value === 'object' || typeof value === 'function') &&
    typeof (value as { then?: unknown }).then === 'function'
  );
}

/**
 * Collaboration, snapshot, formula-result, and transport-only markers are
 * engine provenance, not caller permissions. Remove them at the documented
 * raw-facade escape hatch so a host cannot relabel a local command as
 * authoritative replay, internal recalculation input, or a mutation that the
 * host's persistence stream should ignore. Other Univer execution options
 * remain available to power hosts.
 */
export function sanitizePublicExecutionOptions(
  options: ExecutionOptions | undefined,
): ExecutionOptions | undefined {
  if (!options) return undefined;
  const sanitized = { ...options };
  delete sanitized.fromCollab;
  delete sanitized.fromChangeset;
  delete sanitized.fromFormula;
  delete sanitized.applyFormulaCalculationResult;
  delete sanitized.onlyLocal;
  delete sanitized.syncOnly;
  return sanitized;
}

function wrapPublicCommandListener(callback: PublicCallback): PublicCallback {
  return (command: unknown, options: unknown) =>
    callback(
      clonePlainCommandData(command),
      sanitizePublicExecutionOptions(options as ExecutionOptions | undefined),
    );
}

function wrapPublicCommandEventListener(callback: PublicCallback): PublicCallback {
  return (event: unknown) => {
    if (event === null || typeof event !== 'object') return callback(event);
    const source = event as Record<PropertyKey, unknown>;
    const detached = clonePlainCommandData(source);
    const publicEvent: Record<PropertyKey, unknown> = {
      ...detached,
      options: sanitizePublicExecutionOptions(source.options as ExecutionOptions | undefined),
    };
    const result = callback(publicEvent);
    // BeforeCommandExecute intentionally lets hosts cancel an action. Preserve
    // that documented control while keeping engine provenance isolated.
    if (Object.prototype.hasOwnProperty.call(publicEvent, 'cancel')) {
      source.cancel = publicEvent.cancel;
    }
    return result;
  };
}

/**
 * Preserve the full FUniver facade while guarding both public command methods.
 * Methods execute against the real instance so Univer class internals never
 * see the Proxy as `this`; fluent methods that return that instance are mapped
 * back to the guarded facade so chaining cannot escape the policy boundary.
 */
export function createPublicUniverFacade<T extends object>(facade: T): T {
  const boundMethods = new Map<PropertyKey, { source: unknown; bound: unknown }>();
  const publicFacade = new Proxy(facade, {
    get(target, property) {
      const value = Reflect.get(target, property, target);
      if (typeof value !== 'function') return value;
      const cached = boundMethods.get(property);
      if (cached?.source === value) return cached.bound;
      const guarded = (...inputArgs: unknown[]) => {
        const args = [...inputArgs];
        if (PUBLIC_EXECUTION_METHODS.has(property)) {
          args[2] = sanitizePublicExecutionOptions(args[2] as ExecutionOptions | undefined);
        } else if (PUBLIC_COMMAND_LISTENER_METHODS.has(property)) {
          if (typeof args[0] === 'function') {
            args[0] = wrapPublicCommandListener(args[0] as PublicCallback);
          }
        } else if (
          property === 'addEvent' &&
          PUBLIC_COMMAND_EVENTS.has(String(args[0])) &&
          typeof args[1] === 'function'
        ) {
          args[1] = wrapPublicCommandEventListener(args[1] as PublicCallback);
        }
        const result = Reflect.apply(value, target, args);
        if (result === target) return publicFacade;
        if (isPromiseLike(result)) {
          return Promise.resolve(result).then((resolved) =>
            resolved === target ? publicFacade : resolved,
          );
        }
        return result;
      };
      boundMethods.set(property, { source: value, bound: guarded });
      return guarded;
    },
  });
  return publicFacade;
}
