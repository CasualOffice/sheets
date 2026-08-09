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
 * Univer 0.25 pushes a command onto its private execution stack before calling
 * `beforeCommandExecuted`, but its published build does not pop that item when
 * a listener vetoes by throwing `CustomCommandExecutionError`. A stale command
 * is then misreported as the trigger of later mutations.
 *
 * Remove only the exact command-info object Univer just handed to our listener.
 * This is safe for nested execution (parent stack entries remain intact), is a
 * no-op on versions without that implementation detail, and is also compatible
 * with newer Univer builds that independently clean the stack in `finally`.
 */
export function cleanCommandExecutionStackForVeto(
  commandService: unknown,
  commandInfo: object,
): boolean {
  const stack = (commandService as { _commandExecutionStack?: unknown })?._commandExecutionStack;
  if (!Array.isArray(stack)) return false;

  for (let index = stack.length - 1; index >= 0; index -= 1) {
    if (stack[index] !== commandInfo) continue;
    stack.splice(index, 1);
    return true;
  }
  return false;
}
