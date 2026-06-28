#!/usr/bin/env bash
# Copyright 2026 Casual Office
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

# Prepare the vendored Univer fork for consumption by the sheet workspace.
#
# Runs:
#   1. `pnpm install` inside the fork to populate vendor/univer-revamp/node_modules
#   2. `pnpm build` inside the fork to produce lib/ outputs (.d.ts + .js)
#   3. swap-fork-pkgs.mjs to flip each fork package.json's `main` / `exports`
#      from the dev shape (./src/index.ts) to the consumable shape
#      (./lib/es/index.js etc.) — same swap upstream applies at publish time
#      via the publishConfig block.
#
# After this script, the sheet workspace's `pnpm install` resolves
# @univerjs/* via the `pnpm.overrides` block in the root package.json
# and tsc reads the built .d.ts files from lib/types/.
#
# Idempotent: re-running is safe. The swap script is reversible via
# `scripts/swap-fork-pkgs.mjs --restore`.
#
# Caller is responsible for checking out the submodule first
# (`git submodule update --init --recursive` or
# `actions/checkout@v4` with `submodules: recursive`).

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")"/.. && pwd)"
FORK_DIR="$REPO_ROOT/vendor/univer-revamp"

if [[ ! -d "$FORK_DIR" ]]; then
  echo "error: fork submodule not found at $FORK_DIR — run 'git submodule update --init --recursive' first" >&2
  exit 1
fi

echo "==> installing fork dependencies ($FORK_DIR)"
( cd "$FORK_DIR" && pnpm install --frozen-lockfile )

# Always restore the dev shape before building. The build relies on
# cross-package source resolution (`main: ./src/index.ts`); leaving a
# previous run's swap in place breaks turbo's parallel build because
# the lib/ outputs for upstream packages don't exist yet at the moment
# downstream packages try to typecheck against them.
echo "==> restoring fork package.jsons to src/ shape (in case a prior swap is in place)"
node "$REPO_ROOT/scripts/swap-fork-pkgs.mjs" --restore

echo "==> building fork"
# The fork build is a big parallel turbo run (~50 packages). On a loaded CI
# runner it can OOM/flake on a single package (a different one each time —
# ui-adapter-vue3, sheets-find-replace, …), failing the whole job for no real
# reason. turbo caches successful packages, so a retry only rebuilds the ones
# that failed → fast and reliable. Retry up to 3x before giving up.
fork_build() { ( cd "$FORK_DIR" && pnpm build ); }
attempt=1
until fork_build; do
  if [ "$attempt" -ge 3 ]; then
    echo "==> fork build failed after $attempt attempts" >&2
    exit 1
  fi
  echo "==> fork build attempt $attempt failed (likely a transient runner OOM); retrying…" >&2
  attempt=$((attempt + 1))
done

echo "==> swapping fork package.json main/exports to use built lib/"
node "$REPO_ROOT/scripts/swap-fork-pkgs.mjs"

echo "==> fork ready; run 'pnpm install' at the repo root next."
