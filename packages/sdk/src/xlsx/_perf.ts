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
 * Tiny perf shim — drop-in replacement for apps/web's `../perf` so the
 * SDK doesn't have to pull in a User Timing dependency. The host can
 * wrap individual SDK calls with its own timing if it wants spans;
 * inside the SDK they're cheap no-op forwards.
 */

export function timeIt<T>(_label: string, fn: () => T): T {
  return fn();
}

export async function timeItAsync<T>(_label: string, fn: () => Promise<T>): Promise<T> {
  return fn();
}
