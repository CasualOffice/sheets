# SKILLS.md — how to do common tasks in this repo

Practical command reference for working on Casual Sheets. For *what* the project is and
*why*, see [`CLAUDE.md`](./CLAUDE.md), [`docs/ARCHITECTURE.md`](./docs/ARCHITECTURE.md),
and [`docs/SDK_ARCHITECTURE.md`](./docs/SDK_ARCHITECTURE.md).

Monorepo: pnpm workspaces (`apps/*`, `packages/*`). Node ≥ 18.17.

---

## Run locally

```bash
pnpm install
pnpm dev:web        # web app (Vite)        → @sheet/web
pnpm dev:server     # collab/storage server → @sheet/server  (only needed for co-editing)
```

The web app runs with no server — `BrowserFileSource` (IndexedDB) is the zero-config
default. Start `dev:server` only when exercising collaboration, WOPI, or personal mode.

---

## Validate before every push  *(required)*

Run all of these and get them green **before** `git push` — not just typecheck + unit:

```bash
pnpm lint            # eslint .
pnpm format:check    # prettier --check .   (CI gates on this)
pnpm typecheck       # pnpm -r typecheck
pnpm test:unit       # node scripts/test-unit.mjs
pnpm build:web       # production build must succeed
```

### UI changes also require Playwright  *(non-negotiable)*

Any change that touches rendered UI must be **driven through Playwright** (observe the
real screen/flow) and pass CI before reaching origin. Typecheck + unit do not catch UI
regressions.

```bash
pnpm test:e2e                                   # default config (playwright.config.ts)
pnpm test:e2e:ui                                # interactive runner
pnpm exec playwright test -c playwright.personal.config.ts   # personal mode
pnpm exec playwright test -c playwright.wopi.config.ts       # WOPI mode
pnpm exec playwright test -c playwright.docker.config.ts     # docker image
```

Work in **small batches (3–4 commits) → push → wait for green CI** before piling on more.

---

## Univer fork

The fork is the git submodule `vendor/univer-revamp` (`schnsrw/univer-revamp`). The app
consumes `@univerjs/*` via `pnpm.overrides` pointing at the fork — every imported
`@univerjs/*` (direct + transitive) must be listed in the overrides block or it resolves
to npm and DI breaks.

```bash
pnpm fork:setup      # ./scripts/setup-fork.sh
pnpm fork:swap       # node scripts/swap-fork-pkgs.mjs           (use fork builds)
pnpm fork:restore    # node scripts/swap-fork-pkgs.mjs --restore (back to npm)
```

Bumping the fork = `git -C vendor/univer-revamp checkout <branch/sha>` + commit on this
repo. The 0.24→0.25 upgrade is **Phase 0** of
[`docs/SDK_MIGRATION_PIPELINE.md`](./docs/SDK_MIGRATION_PIPELINE.md). See also
[`docs/UNIVER_FORK_PERF.md`](./docs/UNIVER_FORK_PERF.md).

---

## Release the SDK

`@casualoffice/sheets` is the published package; releases go through Changesets.

```bash
# 1. add a changeset describing the change (run in repo root)
pnpm dlx @changesets/cli@2 add
# 2. version (consumes changesets, bumps versions, updates CHANGELOG)
pnpm version-packages
# 3. build + publish
pnpm release         # builds @casualoffice/sheets then changesets publish
```

Never break the props / `CasualSheetsAPI` surface without a major changeset — it is the
semver contract integrators depend on.

---

## Docker

```bash
docker compose -f docker-compose.dev.yml up   # local dev stack
docker compose up                             # production-shaped stack
```

Image + self-hosting docs: [`docs/DOCKERHUB.md`](./docs/DOCKERHUB.md),
[`docs/self-hosting/overview.md`](./docs/self-hosting/overview.md).
