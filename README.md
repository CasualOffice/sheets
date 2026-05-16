# Casual Sheets

A web-based, Excel-flavored spreadsheet editor with real-time collaboration on the roadmap. Built on [Univer](https://github.com/dream-num/univer) (Apache-2.0 OSS).

The goal: feel like Excel, not like Google Sheets — ribbon, formula bar, file-centric workflow.

## Status

Phase 1 — single-user editor is feature-rich. Multi-user (Yjs + Hocuspocus) comes next.

| Working | Coming |
| --- | --- |
| Office-style ribbon (Home / Insert / Formulas / Data) | Real-time co-editing |
| Inline cell editing, F2, Backspace, Delete, Escape | Presence (cursors, avatars) |
| Formula bar with Name Box, fx, live mirror | More ribbon tabs (Review / View) |
| Fonts (family + size), colors, fill, wrap, alignment | Charts |
| Borders (split-button dropdown with 7 modes) | Pivot tables |
| Cell merge + unmerge | Print / PDF |
| AutoSum / Average / Count / Min / Max | WOPI host integration |
| Sort ascending/descending, Filter | |
| Open / Save As .xlsx (ExcelJS round-trip) | |
| Sheet tabs at the bottom (add / rename / delete / switch) | |
| Auto-fit column / row | |
| File menu with Properties dialog | |
| Dynamic workbook growth — 1024×128 → 8192×1024 | |
| Material Symbols icons, Inter typography | |

## Develop

Prereqs: Node ≥ 18.17, pnpm 10+.

```sh
pnpm install               # one-time
pnpm dev:web               # http://127.0.0.1:5273
pnpm dev:server            # http://127.0.0.1:3000 (placeholder until collab lands)
pnpm test:e2e              # Playwright (auto-starts the web dev server)
pnpm test:e2e:ui           # Playwright UI mode
pnpm lint                  # eslint
pnpm format                # prettier --write
pnpm typecheck             # tsc across packages
```

## Repo layout

```
.
├── apps/
│   ├── web/                ← Vite + React frontend
│   └── server/             ← Fastify scaffold (Hocuspocus + xlsx routes land in Phase 2)
├── tests/e2e/              ← Playwright e2e suite
├── docs/
│   ├── ARCHITECTURE.md     ← system design
│   └── RESEARCH.md         ← Univer technical brief
├── vendor/univer/          ← read-only Univer 0.22.1 source clone (gitignored)
├── PLAN.md                 ← phased build plan
└── CLAUDE.md               ← project guardrails
```

`vendor/univer/` is a local clone of `dream-num/univer` for source-level study and is excluded from version control. To bootstrap on a fresh checkout:

```sh
git clone --depth 1 https://github.com/dream-num/univer.git vendor/univer
```

## Stack

| Concern | Pick |
| --- | --- |
| Editor + formula engine | Univer OSS (`@univerjs/core` + sheets plugins, pinned to 0.22.1) |
| Frontend | React + Vite + TypeScript (strict) |
| Lint / format | ESLint 9 (flat config) + Prettier |
| xlsx I/O | ExcelJS |
| Icons / type | Material Symbols Outlined + Inter (Google Fonts) |
| E2E tests | Playwright (Chromium) |
| Collab transport (Phase 2) | Yjs + Hocuspocus |

## License

Project code: TBD. Vendored Univer source is Apache-2.0 (DreamNum Co., Ltd.).
