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
 * Detach serializable command payloads before exposing them to host callbacks.
 * Univer commands use plain objects/arrays; opaque class instances are retained
 * by reference because guessing how to clone plugin-owned state is less safe.
 * Cycles are supported defensively.
 */
export function clonePlainCommandData<T>(value: T, seen = new WeakMap<object, unknown>()): T {
  if (value === null || typeof value !== 'object') return value;
  const source = value as object;
  const prior = seen.get(source);
  if (prior !== undefined) return prior as T;

  if (Array.isArray(value)) {
    const result: unknown[] = [];
    seen.set(source, result);
    for (const item of value) result.push(clonePlainCommandData(item, seen));
    return result as T;
  }

  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) return value;

  const result = Object.create(prototype) as Record<PropertyKey, unknown>;
  seen.set(source, result);
  for (const key of Reflect.ownKeys(value)) {
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor?.enumerable) continue;
    result[key] = clonePlainCommandData((value as Record<PropertyKey, unknown>)[key], seen);
  }
  return result as T;
}
