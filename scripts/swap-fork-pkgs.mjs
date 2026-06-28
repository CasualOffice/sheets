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
 * Swap each fork package's `main` / `exports` from the dev-mode shape
 * (`./src/index.ts`) to the consumable shape (`./lib/es/index.js`
 * etc., the values upstream tucks under `publishConfig`).
 *
 * Background: the fork at vendor/univer-revamp/ is wired into the
 * sheet workspace via `pnpm.overrides` + `link:` paths. The fork
 * ships its package.json's `main` and `exports` pointing at
 * `./src/index.ts` for in-fork development; when consumed by the
 * sheet's strict tsconfig, tsc walks the .ts source and trips on
 * the fork's looser type rules (decorator declarations, unused
 * locals, etc.).
 *
 * Published @univerjs/* packages avoid this because npm publish
 * applies `publishConfig` — overrides `main` + `exports` to point
 * at the built `lib/` outputs. We do the same swap locally after
 * the fork builds.
 *
 * Invocation:
 *   node scripts/swap-fork-pkgs.mjs                # apply swap
 *   node scripts/swap-fork-pkgs.mjs --restore     # undo (reverts
 *                                                 to dev shape)
 *
 * Run via `pnpm fork:swap` and `pnpm fork:restore`. Mirrors the
 * standard package.json `publishConfig` semantics.
 */

import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { readdirSync } from 'node:fs';

const here = fileURLToPath(new URL('.', import.meta.url));
const FORK_PACKAGES = join(here, '..', 'vendor', 'univer-revamp', 'packages');
const MARKER = '__casual_sheets_swapped';

const args = new Set(process.argv.slice(2));
const restore = args.has('--restore');

const packageDirs = readdirSync(FORK_PACKAGES).filter((name) => {
  const pkgPath = join(FORK_PACKAGES, name, 'package.json');
  return existsSync(pkgPath);
});

let changed = 0;
let skipped = 0;

for (const name of packageDirs) {
  const pkgPath = join(FORK_PACKAGES, name, 'package.json');
  const raw = readFileSync(pkgPath, 'utf8');
  const pkg = JSON.parse(raw);

  if (restore) {
    if (!pkg[MARKER]) {
      skipped += 1;
      continue;
    }
    // Restore the dev shape from the backup we tucked alongside.
    pkg.main = pkg[MARKER].main;
    pkg.module = pkg[MARKER].module;
    pkg.exports = pkg[MARKER].exports;
    delete pkg[MARKER];
    writeFileSync(pkgPath, JSON.stringify(pkg, null, 4) + '\n', 'utf8');
    changed += 1;
    continue;
  }

  // Apply: only when publishConfig exists and we haven't already swapped.
  if (pkg[MARKER]) {
    skipped += 1;
    continue;
  }
  if (!pkg.publishConfig) {
    skipped += 1;
    continue;
  }
  // Stash the dev shape.
  pkg[MARKER] = {
    main: pkg.main,
    module: pkg.module,
    exports: pkg.exports,
  };
  if (pkg.publishConfig.main) pkg.main = pkg.publishConfig.main;
  if (pkg.publishConfig.module) pkg.module = pkg.publishConfig.module;
  if (pkg.publishConfig.exports) pkg.exports = pkg.publishConfig.exports;
  writeFileSync(pkgPath, JSON.stringify(pkg, null, 4) + '\n', 'utf8');
  changed += 1;
}

const action = restore ? 'restored' : 'swapped';
console.log(`${action} ${changed} fork package(s); skipped ${skipped}`);
