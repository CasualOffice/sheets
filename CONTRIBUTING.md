# Contributing to Casual Sheets

Casual Sheets is a web-based, Excel-equivalent spreadsheet editor built on Univer OSS.
Its primary purpose is to be an **embeddable editor SDK** other apps and engines attach
to — the hosted site is a reference consumer, not the product. Read
[`docs/SDK_ARCHITECTURE.md`](./docs/SDK_ARCHITECTURE.md) before substantial work.

## Setup

```bash
pnpm install
pnpm dev:web      # editor; runs with no server (localStorage default)
pnpm dev:server   # only needed for collaboration / WOPI / personal mode
```

Node ≥ 18.17, pnpm workspaces. Command reference: [`SKILLS.md`](./SKILLS.md).

## The verification gate  *(required before every push)*

1. **Full local validation:** `pnpm lint && pnpm format:check && pnpm typecheck &&
   pnpm test:unit && pnpm build:web` — all green.
2. **UI changes:** must additionally be **driven through Playwright** (run the real app
   and observe the affected flow) and pass CI before reaching origin. Typecheck + unit
   do not catch UI regressions. See `SKILLS.md` for the Playwright configs.
3. **Wait for green CI** on each batch before pushing more.

Work in **small batches** — 3–4 commits, push, confirm CI is green, then continue.
Smaller batches keep red intervals short and root-cause obvious.

## Code conventions

- TypeScript everywhere, strict mode. React + Vite frontend.
- Match the existing tight, decision-oriented tone in `PLAN.md` and `docs/`. No
  marketing language in docs — state decisions and tradeoffs.
- Chrome uses **Inter** (via Google Fonts) and **Material Symbols (Outlined)** for icons
  — never text glyphs or other icon libraries. Adopt `@schnsrw/design-system` primitives
  where available.
- When citing Univer internals, use `vendor/univer-revamp/packages/.../file.ts:LINE`.
- Respect the hard rules in [`CLAUDE.md`](./CLAUDE.md): the collab hook
  (`onMutationExecutedForCollab`), the pinned Univer version, no Univer Pro.

## Commits & PRs

- Conventional-commit prefixes (`feat:`, `fix:`, `chore:`, `docs:`…), matching the
  existing history.
- Any change to the published SDK (`@casualoffice/sheets`) needs a **changeset**
  (`pnpm dlx @changesets/cli@2 add`). Never break the props / `CasualSheetsAPI` surface
  without a major bump — it is the contract integrators depend on.
- Keep PRs scoped to one batch; link the doc/phase they advance
  (`docs/SDK_MIGRATION_PIPELINE.md`).

## Where things live

| Area | Path |
| --- | --- |
| Published SDK | `packages/sdk` (`@casualoffice/sheets`) |
| Web host | `apps/web` |
| Collab/storage server | `apps/server` |
| Univer fork (submodule) | `vendor/univer-revamp` |
| Architecture (current / target) | `docs/ARCHITECTURE.md` / `docs/SDK_ARCHITECTURE.md` |
| Migration plan | `docs/SDK_MIGRATION_PIPELINE.md` |
