# CLAUDE.md — instructions for Claude Code in this repo

## What this project is

A web-based **Excel-equivalent** with real-time collaborative editing, built on **Univer OSS** (Apache-2.0). The goal UX is Microsoft Excel / Office — ribbon, formula bar, file-centric flow — not Google Sheets.

## What's in scope

- Upload `.xlsx` → open in browser session → multi-user co-edit → download `.xlsx`.
- In-memory sessions only. No database. No accounts.
- Office-style UI shell built on top of Univer's grid + formula engine.

## What's out of scope (do not propose, do not build)

> **Note**: parts of this section are out of date — Phase C personal mode, Phase D WOPI, autosave, and version history have all shipped since this list was written. The remaining "do not build" entries below are still binding.

- **AI / LLM features** — the user will plug in a self-hosted LLM later through Univer's command bus. Don't pre-design for it.
- **Mobile** — supported as a **viewer + light editor** down to ~360 px (iPhone SE+). Open files, scroll, single-cell value edits, basic formatting via the menu strip + compact toolbar, sheet switching. NOT supported: chart insert dialogs, pivot field-list, complex formula composition, or any flow that needs hover + right-click on phone. Breakpoints live in `apps/web/src/styles.css` at `@media (max-width: 720px)` and `@media (max-width: 480px)`. iOS Safari requires input font-size ≥ 16 px to avoid focus-zoom; honour that on any input inside the chrome. Univer's canvas owns its own touch gestures — don't try to wrap them.
- **Univer Pro features** — collab, xlsx I/O, charts, pivots, print, history are all paid in Univer's commercial offering. We are *not* using Univer Pro. We build the collab + xlsx layers ourselves on OSS.

## Required reading before substantive work

1. [`PLAN.md`](./PLAN.md) — phased plan and estimates.
2. [`docs/ARCHITECTURE.md`](./docs/ARCHITECTURE.md) — how the pieces fit today.
3. [`docs/SDK_ARCHITECTURE.md`](./docs/SDK_ARCHITECTURE.md) — **target** architecture: this repo's primary purpose is an embeddable editor SDK other engines attach to (the Excalidraw model: package *is* the editor, opt-in collab server, thin localStorage host).
4. [`docs/SDK_MIGRATION_PIPELINE.md`](./docs/SDK_MIGRATION_PIPELINE.md) — phased path to that target (Phase 0 = Univer 0.25).
5. [`docs/RESEARCH.md`](./docs/RESEARCH.md) — Univer technical brief with file path references.
6. [`SKILLS.md`](./SKILLS.md) — build/test/verify/release/fork workflows. [`CONTRIBUTING.md`](./CONTRIBUTING.md) — contribution + verification gate.
7. [`docs/RELEASING.md`](./docs/RELEASING.md) — the **two independent release lines**: Docker app (`casualoffice/sheets`, `vX.Y.Z` tags, latest 0.3.1) vs npm SDK (`@casualoffice/sheets`, Changesets, latest 0.8.0). They version separately — don't conflate them.

## Hard rules

### `vendor/univer/` is our fork — modifiable, not yet wired into the build

- It is a full clone of [`CasualOffice/univer-revamp`](https://github.com/CasualOffice/univer-revamp), our long-term fork of `dream-num/univer` (currently at v0.24.0).
- **You may modify files under `vendor/univer/`** — commits land in the fork.
- **Still not included in our build or workspace.** The app consumes `@univerjs/*` from npm (pinned, see `apps/web/package.json`). Until/unless we wire the fork via pnpm overrides or a republished scope, modifications here are upstream work, not local runtime changes.
- Read it freely. Cite file paths and line numbers when explaining Univer internals; `vendor/univer/packages/.../file.ts:LINE` format still applies.
- Active perf plan lives at [`docs/UNIVER_FORK_PERF.md`](./docs/UNIVER_FORK_PERF.md).

### Pin Univer version

- Per the research brief, Univer's `IWorkbookData` shape and plugin contracts change across minor versions, with strict version validation between plugins.
- Pin **all** `@univerjs/*` packages to the exact same version. Never mix. Current pin: **0.24.0** (`apps/web/package.json` + `packages/sdk/package.json`).
- **0.25 upgrade (Phase 0, in progress):** the vendored submodule `vendor/univer-revamp` (remote `CasualOffice/univer-revamp`) sits on branch `casual-sheets/0.24` with **six** custom commits on top of the `v0.24.0` release — 2 feature (paste-merge preservation, filtered-dropdown visibility) + 4 perf (font-cache LRU, merge-range row-bucket index, header hit-test index, setStylesCache span). No 0.25 exists in the fork. Upstream `dream-num/univer` tags `v0.25.0` (`36a3884c`). Upgrade = branch `casual-sheets/0.25` from upstream `v0.25.0`, cherry-pick the six commits, bump every pin `0.24.0 → 0.25.0`, re-audit the `pnpm.overrides` block, retest. See `docs/SDK_MIGRATION_PIPELINE.md` Phase 0.

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

### Verify UI changes before pushing

- **Every UI change must be driven through Playwright** (run the real app, observe the affected screen/flow) **and pass CI before it reaches origin.** Typecheck + unit alone is never sufficient for UI — regressions don't show up there, and the polished-UX bar requires seeing the rendered result.
- Before any `git push`: run full local validation — `lint` + `format:check` + `typecheck` + `test:unit` + `build`, plus the relevant `test:e2e` Playwright config — then wait for green CI. Work in small batches (3–4 commits) so each verified slice is independently pushable.

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

Original phase plan (kept for historical context):

- ~~**Phase 0**~~ — spikes only. ✅ Done.
- ~~**Phase 1**~~ — single-player editor + Office shell. ✅ Shipped.
- ~~**Phase 2**~~ — Yjs collab. ✅ Shipped via Hocuspocus.
- ~~**Phase 3**~~ — presence. ✅ Shipped (`PresenceLayer` + `AvatarStack`).

Subsequent phases (also shipped):

- **Phase C — Personal mode** (single + multi). bcrypt + SQLite users
  table, per-user file scoping, admin panel. See `docs/self-hosting/
  personal-mode.md`.
- **Phase D — WOPI** (Mode 2). JWT verifier + Lock/Unlock + refresh
  ticker. See `docs/STORAGE_MODES.md`.
- **M2 — Snapshot pipeline.** Client-push autosave via
  `useFileSourceAutoSave` (Bun worker pool deferred — the client
  push covers the practical case).

## Current status (2026-06-12)

- **Released**: v0.3.1 (Docker image + pnpm packages). Single-user
  demo live at https://sheet.casualoffice.org/.
- **Recent UX wave** (UX_AUDIT.md, all shipped 2026-06-11/12): path
  router with `/home` file picker, mobile-responsive list, keyboard
  shortcuts dialog, SaveStatusPill, ActivityPill, collab name pre-
  fill, Ctrl+Shift+P command-palette alias, `formatShortcut` util
  closing the Mac shortcut-rendering debt.
- **Deferred / pending**: sharing-model implementation (design lives
  in `docs/SHARING_MODEL.md`); per-entry retry handlers on
  ActivityPill.
- **CI**: green. Go toolchain pinned 1.25. Smoke + audit specs run
  on every PR.

### Active direction (2026-06-19): SDK-first restructure

The repo's primary purpose is an **embeddable editor SDK** other engines
attach to — Excalidraw's model. Direction locked: promote the full
editor into `@casualoffice/sheets` (today it only boots a minimal
Univer; the real editor lives in `apps/web/src/UniverSheet.tsx` +
`apps/web/src/shell/`), make storage (localStorage default) + collab
opt-in adapters, slim `apps/web` to a thin SDK consumer, adopt
`@schnsrw/design-system`. Sequenced behind a Univer 0.24→0.25 fork
upgrade (Phase 0). Full plan: `docs/SDK_ARCHITECTURE.md` +
`docs/SDK_MIGRATION_PIPELINE.md`.

When in doubt about what's shipped vs. pending, check `docs/UX_AUDIT.md`
§5 (last refreshed 2026-06-12) — it tracks each item with the
commit SHA that shipped it.
