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
 * Adds the Apache-2.0 copyright header to every first-party source file.
 *
 * Scope: tracked .ts/.tsx/.js/.mjs/.css/.sh/.py files under apps/,
 * packages/, scripts/, and tests/. Excludes the vendored Univer fork
 * (vendor/, which keeps DreamNum's upstream headers), generated output,
 * JSON (no comment syntax), and binary/asset files.
 *
 * Idempotent: a file that already carries the header is left untouched,
 * so this is safe to re-run (e.g. in a pre-commit check).
 */
import { execSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { extname } from 'node:path';

const MARKER = 'Copyright 2026 Casual Office';

const LICENSE_LINES = [
  'Copyright 2026 Casual Office',
  '',
  'Licensed under the Apache License, Version 2.0 (the "License");',
  'you may not use this file except in compliance with the License.',
  'You may obtain a copy of the License at',
  '',
  '    http://www.apache.org/licenses/LICENSE-2.0',
  '',
  'Unless required by applicable law or agreed to in writing, software',
  'distributed under the License is distributed on an "AS IS" BASIS,',
  'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.',
  'See the License for the specific language governing permissions and',
  'limitations under the License.',
];

const BLOCK_HEADER = ['/**', ...LICENSE_LINES.map((l) => (l ? ` * ${l}` : ' *')), ' */'].join('\n');
const HASH_HEADER = LICENSE_LINES.map((l) => (l ? `# ${l}` : '#')).join('\n');

const BLOCK_EXTS = new Set(['.ts', '.tsx', '.js', '.mjs', '.cjs', '.css']);
const HASH_EXTS = new Set(['.sh', '.py']);

const tracked = execSync('git ls-files apps packages scripts tests', { encoding: 'utf8' })
  .split('\n')
  .map((s) => s.trim())
  .filter(Boolean)
  .filter((p) => !p.startsWith('vendor/'))
  .filter((p) => !/\/(node_modules|dist|build|lib|\.vite)\//.test(p));

let added = 0;
let skipped = 0;

for (const file of tracked) {
  const ext = extname(file);
  const block = BLOCK_EXTS.has(ext);
  const hash = HASH_EXTS.has(ext);
  if (!block && !hash) continue;

  const content = readFileSync(file, 'utf8');
  if (content.includes(MARKER)) {
    skipped += 1;
    continue;
  }

  const header = block ? BLOCK_HEADER : HASH_HEADER;
  const lines = content.split('\n');
  const hasShebang = lines[0]?.startsWith('#!');
  const shebang = hasShebang ? `${lines[0]}\n` : '';
  const rest = (hasShebang ? lines.slice(1).join('\n') : content).replace(/^\n+/, '');

  writeFileSync(file, `${shebang}${header}\n\n${rest}`);
  added += 1;
}

console.log(`[license-headers] added ${added}, skipped ${skipped} (already had header)`);
