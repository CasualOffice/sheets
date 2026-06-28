#!/usr/bin/env node
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
 * Discover *.unit.test.ts files under apps/ and packages/, then spawn
 * `node --import tsx --test` on the collected list. Node 20's test
 * runner doesn't expand glob patterns itself; doing it here keeps the
 * `pnpm test:unit` script portable without adding a glob dependency.
 */
import { spawnSync } from 'node:child_process';
import { readdirSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const ROOTS = ['apps', 'packages'];
const SKIP_DIRS = new Set(['node_modules', 'dist', 'build', '.next', 'lib', '.turbo', 'coverage']);

function walk(dir, acc) {
  let entries;
  try {
    entries = readdirSync(dir);
  } catch {
    return;
  }
  for (const entry of entries) {
    if (SKIP_DIRS.has(entry)) continue;
    const full = join(dir, entry);
    let st;
    try {
      st = statSync(full);
    } catch {
      continue;
    }
    if (st.isDirectory()) walk(full, acc);
    else if (entry.endsWith('.unit.test.ts') || entry.endsWith('.unit.test.tsx')) acc.push(full);
  }
}

const files = [];
for (const r of ROOTS) walk(join(ROOT, r), files);

if (files.length === 0) {
  console.log('[test:unit] no *.unit.test.ts files found — skipping');
  process.exit(0);
}

console.log('[test:unit] running', files.length, 'test file(s)');
const result = spawnSync('node', ['--import', 'tsx', '--test', ...files], {
  stdio: 'inherit',
  cwd: ROOT,
});
process.exit(result.status ?? 1);
