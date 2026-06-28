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
 * Performance marks. Wraps `performance.mark` + `performance.measure` so
 * we can see hot paths (parse, mount, save) in the DevTools User Timing
 * track. Marks survive in production builds — the overhead of one mark
 * pair is sub-microsecond, but the visibility pays for itself when
 * profiling on real workbooks.
 *
 * Usage:
 *   timeIt('mount-unit', () => univer.createUnit(...));
 *   await timeItAsync('parse-xlsx', () => xlsxToWorkbookData(buf));
 *
 * Each call also drops a counter into `globalThis.__perf__` so the
 * profiling test harness can read totals without scraping the
 * PerformanceObserver buffer.
 */

type PerfRecord = { label: string; ms: number; at: number };

declare global {
  // eslint-disable-next-line no-var
  var __perf__: PerfRecord[] | undefined;
}

function record(label: string, ms: number): void {
  const buf = (globalThis.__perf__ ??= []);
  buf.push({ label, ms, at: Date.now() });
  // Cap to the last 500 records — we don't want a long session to
  // memory-leak this buffer in dev.
  if (buf.length > 500) buf.splice(0, buf.length - 500);
}

function safeMark(name: string): void {
  try {
    performance.mark(name);
  } catch {
    /* SecurityError in some sandboxed iframes */
  }
}

function safeMeasure(label: string, start: string, end: string): number {
  try {
    const m = performance.measure(label, start, end);
    return m.duration;
  } catch {
    return 0;
  }
}

export function timeIt<T>(label: string, fn: () => T): T {
  const start = `${label}-start-${Math.random().toString(36).slice(2, 7)}`;
  const end = `${label}-end`;
  safeMark(start);
  try {
    return fn();
  } finally {
    safeMark(end);
    record(label, safeMeasure(label, start, end));
  }
}

export async function timeItAsync<T>(label: string, fn: () => Promise<T>): Promise<T> {
  const start = `${label}-start-${Math.random().toString(36).slice(2, 7)}`;
  const end = `${label}-end`;
  safeMark(start);
  try {
    return await fn();
  } finally {
    safeMark(end);
    record(label, safeMeasure(label, start, end));
  }
}

/** For tests: pull the last N records for a label. */
export function perfRecords(label?: string): PerfRecord[] {
  const all = globalThis.__perf__ ?? [];
  return label ? all.filter((r) => r.label === label) : all;
}
