# Competitive Roadmap — closing the gap to Google Sheets / LibreOffice

> Status: **active** (opened 2026-06-23). Companion to [`UX_AUDIT.md`](./UX_AUDIT.md)
> (per-item shipped/pending tracking) and [`UNIVER_FORK_PERF.md`](./UNIVER_FORK_PERF.md)
> (the perf work this roadmap's Phase 2 sequences). Source analysis is code-grounded;
> every claim below cites the file it was verified against.

## 1. Where we stand

Three reference points, three different things:

| Product | What it is | Real threat on |
| --- | --- | --- |
| **Apache OpenOffice Calc** | Native desktop, effectively unmaintained | nothing — we beat it on web, collab, UX |
| **LibreOffice Calc** | Native desktop, actively developed; web/collab only via the separate Collabora Online server | feature depth, native perf, offline |
| **Google Sheets** | Web-native, gold standard for collaborative UX | collab UX, polish, scale, ecosystem |

We win, today, on a combination none of them offer: **Excel-fidelity UX + lossless `.xlsx`
round-trip + modern real-time collab + self-hosted, no-account-lock-in.** This roadmap
closes the gaps without trading that combination away.

### 1.1 UX (our primary axis)

Strong / at parity: Excel-style ribbon + 7-menu bar, formula bar with autocomplete +
live range-picker, stats status bar, 50+ Mac-aware shortcuts, command palette, dark mode,
real-time peer cursors + presence (`apps/web/src/shell/*`, `apps/web/src/collab/`).

Lacking vs Google Sheets:
- **Sharing/permissions UX** — only anonymous room URLs + password gate today
  (`apps/web/src/collab/`). No viewer/commenter/editor roles. → Phase 3 (hybrid: link-roles
  here, full RBAC in Casual Drive).
- **Comments** — threaded + collab-synced, but no @mentions / notifications / resolve
  workflow. → Phase 3.
- **Mobile** — explicitly desktop-first; shrunk chrome + bottom bar, no native apps
  (`apps/web/src/styles.css` @720/@480). → Phase 6.

### 1.2 Performance

Model: canvas render (Univer) with implicit windowing; parse + export + formula all
worker-offloaded; server snapshot cache for <1s co-editor join; 4 shipped fork perf wins.

The headline gap — **interactive grid ceiling**:

| | Interactive max rows | Max cols | Note |
| --- | --- | --- | --- |
| **Casual Sheets** | was 8,192 → **1,048,576** | was 1,024 → **16,384** | raised to Excel parity in Phase 1 (`snapshot.ts`); grows on demand via `useWorkbookGrowth.ts` |
| Google Sheets | unbounded | 18,278 | 10M cells total |
| LibreOffice Calc | 1,048,576 | 16,384 | |
| Excel | 1,048,576 | 16,384 | |

> **Measured (Phase 1):** the feared "O(rowCount) catastrophe" did not materialize.
> Declaring a full 1,048,576-row grid + editing the far edge is a one-time ~170 ms;
> a **pure cell edit at 1M rows is ~1 ms** (cell edits don't rebuild the skeleton).
> The only costly op is **insert/delete-row** at extreme sizes (~517 ms @1M) — owned
> by Phase 2 T2.1 (sparse insert/delete). So the ceiling was raised to parity directly
> and **the planned T1.2 lazy-allocation fork surgery was dropped as unnecessary.**

Nuance (verified, not the scare-version): **imports are not clamped** — `parse-impl.ts:377`
sets `rowCount` to the real data extent, so a large `.xlsx`/CSV opens (subject to the 100 MB
cap and Univer's up-front row/col metadata allocation). The 8,192 cap bites on **blank or
paste-built sheets**, which wall at row 8,192 (`useWorkbookGrowth.ts:59`). The cap was set low
*because* Univer materializes row/column metadata up front (`snapshot.ts:11-19`) — but the
Phase-1 measurement (see §1.2 table note) showed that allocation is cheap enough to raise the
ceiling straight to 1,048,576 × 16,384. The residual cost is insert/delete-row at extreme
sizes, which Phase 2 (T2.1) hardens. Phase 1 still precedes everything else: nothing is worth
building on a grid that walls at 8,192 rows.

Other realities: client-bound scale (Google offloads to server infra); ~500 active collab
docs / ~1,500 clients per process, shardable (`docs/CAPACITY_MODEL.md`); native desktop wins
raw local compute on huge files.

### 1.3 Features

Our strongest axis. Shipped & competitive: 536 functions, charts (13 types + dual-axis +
trendlines + sparklines), pivots (rows/values + filters + drill-down), conditional
formatting, data validation, autofilter + custom sort, full cell formatting incl. rotation,
threaded comments, hyperlinks, images, freeze panes, named ranges, lossless xlsx round-trip
(54/54 audit), ODS/CSV/PDF, collab, version history, autosave, Flash Fill, Goal Seek,
outline/grouping, paste-special.

Lacking vs the field: scripting/macros that *execute* (we passthrough VBA bytes, run nothing);
pivot field-list drag UI; named cell styles; sheet/cell protection; full in-cell rich text;
dynamic/array formulas (basic only); external/connected data; add-on marketplace; AI/Explore
(deliberately deferred per `CLAUDE.md`).

## 2. Principles

1. **Foundation-first, non-overriding.** Each phase builds on the last; no later phase
   forces rework of an earlier one. Scale → harden → collaborate → feature-depth → automate
   → mobile.
2. **Verify before claiming.** Every gap is checked in code before it becomes a tracker.
   (This doc already corrected one false "files won't open" claim — see §1.2.)
3. **Don't trade the moat.** Excel-fidelity, lossless xlsx, self-hosted collab stay intact.
4. **Drive owns identity.** Per the Casual Drive plan, full RBAC/workspaces/invites live in
   Drive; this repo ships only the thin link-role layer (Phase 3).
5. **No AI in scope.** The self-hosted LLM command-bus hook is out of scope here.

## 3. Phases & milestones

Each phase is a GitHub **milestone**; each phase has one **epic tracker** issue carrying its
task checklist. Work proceeds on dedicated branches with PRs and green CI per
`CONTRIBUTING.md`.

### Phase 1 — Grid Scale & Capacity  *(foundation; ✅ complete)*

> **Outcome:** ceiling raised to Excel parity 1,048,576 × 16,384 (#115); paste-to-fit
> growth (#116); 100k-row open validated ~0.85 s (#117); selection stats clamped to
> the used range so full-column stats work at scale (#118). T1.2 (lazy allocation)
> dropped — measured unnecessary.

Goal: the grid holds real-world data. Interactive ceiling toward Excel parity, enabled by
lazy metadata so boot/growth stay fast.

- **T1.1** Raise the interactive ceiling to **1,048,576 × 16,384** (Excel parity).
  ✅ shipped — measurement showed the declared-grid cost is a one-time ~170 ms and pure
  edits don't rebuild the skeleton, so parity was reached directly. (`snapshot.ts`,
  `useWorkbookGrowth.ts`)
- ~~**T1.2** Lazy row/column metadata allocation (fork perf item 7)~~ — **dropped.** The
  Phase-1 measurement (declared-grid rebuild ~170 ms one-time; pure edit ~1 ms at 1M)
  showed the up-front allocation is not a bottleneck. The only costly path, insert/delete-row
  at extreme sizes, is owned by T2.1. Avoids weeks of risky fork surgery.
- **T1.3** Large-paste / large-fill growth correctness — pasting / filling N > cap rows must
  extend the sheet, not truncate. (verify against `useWorkbookGrowth.ts` cap + paste command)
- **T1.4** Large-file open validation — 100k+ row `.xlsx`/CSV fixtures; boot-time budget;
  raise/scale guards. Add large-file e2e. (`apps/web/src/xlsx`, `apps/web/src/ods`)
- **T1.5** Selection stats + Ctrl+A at scale — keep stats cell-cap honest so big selections
  don't freeze. (`apps/web/src/shell/StatusBar.tsx`, `docs/ARCHITECTURE.md:204`)

### Phase 2 — Performance Hardening  *(foundation; ✅ complete)*

> **Outcome:** T2.1 zebra-refresh guard shipped (#119) — insert-row @1M 241→158 ms,
> profile-found. T2.2/T2.4 measured already-handled by Univer 0.25 (`getDiscreteRanges`
> consolidation; bulk edits fire one mutation). T2.3 (incremental dep-tree) investigated
> → no-go: single-edit cost is floored by an intentional 100 ms recalc debounce, leaving
> only ~66 ms behind it — not worth the highest corruption risk. T2.5 (stages 5–6) already
> shipped (hyperlinks-in-snapshot + op-log compaction). **Net: 1 real win + 4 risk-avoidances** —
> the pre-0.25 fork-perf items were largely already addressed; measuring first avoided weeks
> of high-risk fork surgery for marginal gain.

Goal: make the now-bigger grid fast. Sequences the deferred items in `UNIVER_FORK_PERF.md`.

- **T2.1** Sparse row/column insert/delete via shift-offset map (fork item 6).
- **T2.2** Formula dirty-range coalescing (fork item 9, high-risk) — merge dirty rects before
  dependency search.
- **T2.3** Incremental dependency-tree recalc (fork item 10, high-risk; after T2.2).
- **T2.4** Mutation listener fan-out batching (fork item 8).
- **T2.5** Large-file pipeline stages 5–6 — hyperlinks-in-snapshot + Yjs op-log compaction
  (`docs/LARGE_FILE_PIPELINE.md`).

### Phase 3 — Collaboration Depth  *(T3.1/T3.3 unblocked + shipped; T3.2 assignment = fork work)*

Goal: close the Google collaboration-polish gap. Builds on existing collab + presence.

- **T3.4** Hybrid sharing — link roles (viewer / commenter / editor). ✅ **shipped** (#121,
  #128): `applyCommentOnly` engine-layer veto (cells locked, comments work) wired into
  CollabDriver; anonymous `?role=comment` works end-to-end; the share dialog's 3-way role
  picker (Edit / Comment / View) creates comment-only links (#128). Server-token comment
  enforcement (per-mutation filtering) remains a follow-up.
- **T3.2** Comments — resolve / reopen. ✅ **shipped** (#123 resolve, #124 reopen): panel
  Resolve button + a "Resolved" section (read from `SheetsThreadCommentModel`) with reopen.
  **Assignment** remains — needs a new `assignee` field on `IThreadComment` (a fork model
  change), so it's deferred to fork work.
- **T3.1** Comments — authorship + @mentions. ✅ **shipped** (#129 authorship, #130 mentions),
  via a **block-avoiding** approach that never touches `setCurrentUser`. The original slice-1
  (#122) populated `UserManagerService` and broke collab cell-sync: `currentUser$` feeds
  Univer's permission layer (`sheets-ui/menu/permission-menu-util.ts`), so a custom
  current-user id makes peers treat the client as a non-editor and stop applying grid mutations
  (reverted, #125). Instead authorship now rides an out-of-band `casual-comment-authors` Y.Map
  keyed by comment id (stamped from the local presence identity on the add-comment _command_,
  which runs only on the author's client) — exactly how charts sync alongside the op-log (#129).
  @mentions list real collaborators via a host-pluggable `IMentionIOService` override
  (`CasualMentionIOService`) fed by presence peers, with `docs-mention-ui`'s picker registered
  in the comment lazy-group (#130). Precise cross-user identity stays with Casual Drive.
- **T3.3** Comments — @mention surfacing. ✅ **shipped** (#131): a comment (or reply) that
  @-references the local display name gets a "@You" badge + accent in the comments panel
  (matched by name, the shared cross-peer identity). Cross-user **delivery** (email/push) is
  deferred to Casual Drive (identity ownership); a "mentioning me" filter toggle is a follow-up.

### Phase 4 — Feature Depth

Goal: Excel/Sheets parity on the features users notice missing.

- **T4.1** Dynamic / array formulas with spill ranges. ✅ **shipped (verified)** — the 0.25
  fork's formula engine already evaluates dynamic-array functions (SEQUENCE, UNIQUE, FILTER,
  SORT, TRANSPOSE, XLOOKUP…) and spills automatically: the anchor holds the formula, spilled
  cells hold values, and a blocked spill yields `#SPILL!`. Was untested; e2e coverage added
  (`dynamic-arrays.spec.ts`). Remaining polish: the Excel spill-range outline when the anchor
  is selected.
- **T4.2** In-cell rich text (mixed bold/italic/color within one cell).
- **T4.3** Named cell styles (managed Normal/Good/Bad/Heading… that round-trip through xlsx).
- **T4.4** Sheet & cell-range protection (locked cells, protected ranges) beyond collab
  view-only. ⏳ **in progress** — range protection (#133: Data → Protect range / Remove
  range protection) **and** per-sheet protection (Data → Protect sheet) shipped, app-side over
  the worksheet-permission facade. Model: **collab protection** (the protector keeps editing,
  other editors are blocked) — chosen over Excel's block-everyone; the workbook "Make read-only"
  toggle remains the block-everyone option. Follow-ups: a Protect-Sheet dialog with granular
  permission points + password, and xlsx round-trip of protection metadata.

### Phase 5 — Automation / Scripting

Goal: an automation story (non-AI). Leverages the command bus we already broadcast on.

- **T5.1** Macro recorder — capture command-bus mutations into a named macro.
- **T5.2** Macro runner — replay; bind to button / shortcut.
- **T5.3** (stretch) documented scripting API surface for hosts.

### Phase 6 — Mobile

Goal: a real mobile editor, not shrunk chrome. (per `UX_AUDIT.md` §2.15)

- **T6.1** Mobile IA — path router list + editor routes.
- **T6.2** Touch-first editing — single-cell value edit, compact formatting strip, sheet
  switching.
- **T6.3** Responsive chrome polish ≤ 480 px.

## 4. Sequencing rationale

Phase 1 unblocks everything (no point optimizing or adding features to a grid that walls at
8,192 rows). Phase 2 makes Phase 1 usable at scale. Phases 3–5 are independent features that
sit on the stable foundation and can interleave once 1–2 land. Phase 6 is polish over a
finished desktop surface. Within a phase, high-risk recalc work (T2.2 → T2.3) is strictly
ordered.

## 5. Tracker index

| Phase | Milestone | Epic | Status |
| --- | --- | --- | --- |
| 1 — Grid Scale | Milestone 1 | #109 | ✅ complete |
| 2 — Perf Hardening | Milestone 2 | #110 | ✅ complete |
| 3 — Collab Depth | Milestone 3 | #111 | T3.4 + T3.2(resolve/reopen) + T3.1(authorship/mentions) + T3.3(surfacing) shipped; T3.2 assignment = fork work; cross-user delivery → Drive |
| 4 — Feature Depth | Milestone 4 | #112 | planned |
| 5 — Automation | Milestone 5 | #113 | planned |
| 6 — Mobile | Milestone 6 | #114 | planned |

Each epic carries its task checklist (T*.*). Task issues are promoted from the
checklist as work starts.
