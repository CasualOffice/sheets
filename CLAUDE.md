# CLAUDE.md — instructions for Claude Code in this repo

## What this project is

A web-based **Excel-equivalent** with real-time collaborative editing, built on **Univer OSS** (Apache-2.0). The goal UX is Microsoft Excel / Office — ribbon, formula bar, file-centric flow — not Google Sheets.

## What's in scope

- Upload `.xlsx` → open in browser session → multi-user co-edit → download `.xlsx`.
- In-memory sessions only. No database. No accounts.
- Office-style UI shell built on top of Univer's grid + formula engine.

## What's out of scope (do not propose, do not build)

- **Persistence / WOPI** — explicitly deferred. Don't add Postgres, S3, autosave, or version history.
- **AI / LLM features** — the user will plug in a self-hosted LLM later through Univer's command bus. Don't pre-design for it.
- **Auth / sharing UI / permissions model** — anonymous sessions by room URL.
- **Mobile** — desktop browsers only.
- **Univer Pro features** — collab, xlsx I/O, charts, pivots, print, history are all paid in Univer's commercial offering. We are *not* using Univer Pro. We build the collab + xlsx layers ourselves on OSS.

## Required reading before substantive work

1. [`PLAN.md`](./PLAN.md) — phased plan and estimates.
2. [`docs/ARCHITECTURE.md`](./docs/ARCHITECTURE.md) — how the pieces fit.
3. [`docs/RESEARCH.md`](./docs/RESEARCH.md) — Univer technical brief with file path references.

## Hard rules

### `vendor/univer/` is our fork — modifiable, not yet wired into the build

- It is a full clone of [`schnsrw/univer-revamp`](https://github.com/schnsrw/univer-revamp), our long-term fork of `dream-num/univer` (currently at v0.24.0).
- **You may modify files under `vendor/univer/`** — commits land in the fork.
- **Still not included in our build or workspace.** The app consumes `@univerjs/*` from npm (pinned, see `apps/web/package.json`). Until/unless we wire the fork via pnpm overrides or a republished scope, modifications here are upstream work, not local runtime changes.
- Read it freely. Cite file paths and line numbers when explaining Univer internals; `vendor/univer/packages/.../file.ts:LINE` format still applies.
- Active perf plan lives at [`docs/UNIVER_FORK_PERF.md`](./docs/UNIVER_FORK_PERF.md).

### Pin Univer version

- Per the research brief, Univer's `IWorkbookData` shape and plugin contracts change across minor versions, with strict version validation between plugins.
- Pick one version (start: latest stable on npm, currently 0.22.x line) and pin **all** `@univerjs/*` packages to the exact same version. Never mix.

### Use the collab hook Univer designed for it

- Do **not** hook into UI events to capture changes.
- Subscribe to `ICommandService.onMutationExecutedForCollab` (`vendor/univer/packages/core/src/services/command/command.service.ts:404`). This is the only correct hook — it fires for `CommandType.MUTATION` only and includes `syncOnly` mutations.
- Use `IExecutionOptions.fromCollab` when applying remote mutations so they don't re-broadcast (echo loop prevention).
- Respect `params.__splitChunk__` for large mutations (paste large range, copy worksheet).

### Headless seeding caveats

- Server-side xlsx → snapshot conversion runs in Node. Use the plugin set from `vendor/univer/examples/src/node/sdk/index.ts:42` as the known-safe Node baseline.
- Formula evaluation may be async when offloaded to a worker — if seeding, await calc before snapshotting.

### Don't reach for Univer Pro

If you find a feature is missing (charts, pivots, xlsx import/export), the answer is **build it on OSS or defer it**, not "use the Pro package." Pro is closed-source and out of scope.

## Stack conventions (once code starts)

- TypeScript everywhere, strict mode.
- React + Vite for the frontend.
- Hocuspocus + Yjs for the collab server. Stand-alone Node service.
- ExcelJS for xlsx I/O. (SheetJS Community has license caveats — prefer ExcelJS.)
- Tailwind or vanilla CSS — decide before starting Phase 1, document in `docs/ARCHITECTURE.md`.
- Fluent UI icons to match Office look.

## Style

- Match the existing tight, decision-oriented tone of `PLAN.md` and the docs/ files when adding to them.
- Don't bloat docs with marketing language. State decisions and tradeoffs.
- When citing Univer source, use `vendor/univer/packages/.../file.ts:LINE` format.

## Phase awareness

Always know which phase we're in before writing code:

- **Phase 0** (current) — spikes only. Throwaway code that proves a single risk.
- **Phase 1** — single-player editor + Office shell.
- **Phase 2** — Yjs collab.
- **Phase 3** — presence.

Don't start Phase 1 code until Phase 0 spikes are decided.
