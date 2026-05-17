# Casual Sheets

[![CI](https://github.com/schnsrw/sheets/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/schnsrw/sheets/actions/workflows/ci.yml)
[![Deploy](https://github.com/schnsrw/sheets/actions/workflows/deploy-pages.yml/badge.svg?branch=main)](https://github.com/schnsrw/sheets/actions/workflows/deploy-pages.yml)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](./LICENSE)
[![Node](https://img.shields.io/badge/node-%E2%89%A518.17-brightgreen)](#develop)
[![E2E tests](https://img.shields.io/badge/e2e-120%20passing-brightgreen)](./tests/e2e)

**Live demo: <https://sheet.schnsrw.live/>** — auto-deployed from `main` on every push (CI → Pages).

A web-based, Excel-flavored spreadsheet editor with real-time collaboration on the roadmap. Built on [Univer](https://github.com/dream-num/univer) (Apache-2.0 OSS).

The goal: feel like Excel, not like Google Sheets — ribbon, formula bar, file-centric workflow.

## Status

Phase 1 — single-user editor is feature-rich and locked down by 120 Playwright tests. Multi-user (Yjs + Hocuspocus) comes next.

- **CI** runs `lint`, `typecheck`, and the full Playwright suite on every PR and `main` push ([`.github/workflows/ci.yml`](./.github/workflows/ci.yml)).
- **Deploy** publishes `apps/web` to GitHub Pages on every `main` push ([`.github/workflows/deploy-pages.yml`](./.github/workflows/deploy-pages.yml)).
- **Tests**: 120 e2e specs in [`tests/e2e/`](./tests/e2e); xlsx / ods / csv / tsv round-trip, hyperlinks, drag-fill, drag-and-drop file open, formula evaluation, tables, freeze, comments, and more.
- **Live URL**: <https://sheet.schnsrw.live/>

| Working | Coming |
| --- | --- |
| Office-style ribbon (Home / Insert / Formulas / Data) | Real-time co-editing |
| Inline cell editing, F2, Backspace, Delete, Escape | Presence (cursors, avatars) |
| Formula bar with editable Name Box (type `B5` to jump) | More ribbon tabs (Review / View) |
| Fonts (family + size), colors, fill, wrap, alignment | Charts |
| Borders (split-button dropdown with 7 modes) | Pivot tables |
| Cell merge + unmerge | Recent-files / landing page |
| Group / outline rows + columns (collapse / expand) | WOPI host integration |
| AutoSum / Average / Count / Min / Max | |
| Sort ascending/descending, Filter | |
| Help → Report a bug (prefills env to GitHub issue form) | |
| Styled tooltips on every toolbar / shell button | |
| Open / Save .xlsx / .ods / .csv / .tsv (full round-trip) | |
| Drag-and-drop file open + Save / Export toasts | |
| Hyperlinks (import + export round-trip) | |
| Drag-fill handle, Ctrl+D / Ctrl+R, relative-ref formula extension | |
| Tables (Format as Table, themes, Tables panel) | |
| Comments (with corner indicator markers) | |
| Freeze panes (top row / first column / at selection) | |
| Sheet tabs at the bottom (add / rename / delete / reorder) | |
| Auto-fit column / row | |
| Print active sheet (Ctrl+P, headless render to iframe) | |
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

Apache-2.0. See [`LICENSE`](./LICENSE) and [`NOTICE`](./NOTICE) for attribution.

Vendored Univer source (`vendor/univer/`) keeps its upstream Apache-2.0 license; we do not modify it and it is not part of our build.
