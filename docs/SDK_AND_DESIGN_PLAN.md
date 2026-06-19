# SDK Restructure + Design-System Redesign — Combined Plan

Two large tracks are queued against the same codebase: the **SDK restructure**
(move the full editor into `@casualoffice/sheets`, slim `apps/web` onto it — see
[`SDK_MIGRATION_PIPELINE.md`](./SDK_MIGRATION_PIPELINE.md)) and the
**design-system redesign** (adopt `@schnsrw/design-system` tokens + primitives
across the Office chrome). Both rewrite the shell. This doc states the plan for
each and — the load-bearing decision — **how to sequence them so each chrome file
is touched once, not twice.**

**Direction locked:** Univer **0.25 first** (Phase 0, blocking) · SDK infra and
design tokens run **in parallel** · chrome is a **single convergence pass** that
moves shell into the SDK *and* re-skins it with the design system in one edit.

---

## 1. Executive summary

**Track A — SDK restructure.** The SDK editor
(`packages/sdk/src/sheets/CasualSheets.tsx`) is today a bare, *non-computing*
grid: `notExecuteFormula:true` with no RPC plugin and no worker, feature plugins
dropped, no snapshot swap, no chrome. The real editor lives in
`apps/web/src/UniverSheet.tsx` (+ `apps/web/src/univer/*` + ~70 files under
`apps/web/src/shell/`). The work is to lift formula compute, the lazy-plugin
pipeline, the imperative `CasualSheetsAPI`, and eventually the chrome into the
SDK, leaving `apps/web` a thin host. The crux is a DI blocker in Batch 2 (the
`IRPCChannelService` throw, §3) and a worker-in-a-published-library build problem;
chrome extraction is gated on cutting four import seams (collab ×7, `useUniverAPI`
×21, plus FileSource/auth/router in `TitleBar`/`file-actions`).

**Track B — design-system redesign.** `@schnsrw/design-system` `0.1.0` was
*extracted from* the sheet's own `styles.css`, so its tokens already match the app
1:1 (`--color-accent: #0e7490`, `--titlebar-h: 64px`, etc.). Adoption is largely a
dedupe plus swapping ad-hoc chrome primitives for the package's `Button`/`Icon`/
`Pill`/`Avatar`/`Tabs`/etc. Two prior adoption attempts were reverted within
minutes (§7) — both foundered on the same structural fault: **the package is not
published to npm**, so every dependency expression (`link:`, `^0.1.0` + override,
`file:` tarball) worked on the author's machine and broke on CI / fresh clone / the
Drive consumer tarball. The redesign must not repeat that.

**Headline recommendation — sequence by surface, not by track.** Do **Phase 0
(0.25 fork) first; it gates everything.** Then run **pure-SDK-infra work (Track A
Batches 1–3: plugins, DI/formula, the imperative API)** and **pure-design-token
work (Track B token dedupe + low-coupled primitive restyles in `apps/web`)** *in
parallel* — they touch disjoint files. Both tracks then **converge on the chrome**:
extract `apps/web/src/shell/*` into the SDK as design-system-based slot components
in **one pass per file**, so a `TitleBar`/`Toolbar`/`Dialog` is moved-and-reskinned
together. Touching each chrome file twice (once to restyle, once to relocate) is the
waste this sequencing exists to avoid.

---

## 2. The sequencing decision

**Question:** run the SDK restructure and the design redesign simultaneously, or
serialize them?

**Decision: partially parallel, with a forced convergence on the chrome.**

The two tracks decompose into three buckets by *which files they edit*:

| Bucket | Track | Files | Conflict with other track? |
|---|---|---|---|
| SDK infra | A (Batches 1–3) | `packages/sdk/src/univer/*`, `plugins.ts`, `formula-worker.ts`, `CasualSheets.tsx`, the `CasualSheetsAPI` | **No.** No rendered chrome. |
| Design tokens + primitives | B (token dedupe + restyle) | `apps/web/src/styles.css` tokens, `apps/web/src/shell/Icon.tsx`, `Dialog.tsx`, `Popover.tsx`, `Tooltip.tsx`, `RibbonControls.tsx`, status pills | **No** (with one caveat below). |
| Chrome | A (Batch 4) **and** B (final restyle of `shell/*`) | `apps/web/src/shell/*` → `packages/sdk/src/chrome/*` | **Yes — both rewrite the same files.** |

**Why parallel for the first two buckets.** Track A Batches 1–3 are pure Univer
plumbing — formula compute, plugin registration, snapshot swap, the imperative API.
They render no chrome and import none of `shell/*`. Track B's token work edits
`styles.css:5-160` and the presentational primitives (`Icon`/`Dialog`/`Popover`/
`Tooltip`/`RibbonControls`/pills), which are pure-CSS/SVG and read no Univer plumbing.
The file sets are disjoint; running them on separate branches is safe and roughly
halves wall-clock.

**Why convergence is mandatory for the chrome.** Both the SDK chrome-extraction
(Track A Batch 4: `shell/*` → `@casualoffice/sheets/chrome` as slot components) and
the design redesign of the chrome (Track B: restyle `TitleBar`/`MenuBar`/`Toolbar`/
`SheetTabs`/`FormulaBar` with design-system primitives) rewrite **the same ~70 shell
files**. If they run independently, every shell file is edited twice — once to swap
in DS primitives, once to invert its app-coupling and move it — doubling merge
surface, Playwright runs, and review. The thesis holds: **the chrome is the single
convergence point.** Extract each shell file into the SDK *as a design-system-based
slot component* in one edit. Move-and-reskin together.

**The one caveat on parallel-safety:** Track B's token dedupe and the chrome
convergence both eventually own `styles.css`. Keep Track B's parallel phase to
*tokens + presentational primitives only*; do **not** start restyling
`TitleBar`/`MenuBar` in `apps/web` in parallel, because those files are about to be
moved-and-reskinned in the convergence pass. Restyling them in place first is
exactly the double-touch we're avoiding.

**Ordering in one line:** `Phase 0 → (A1–3 ∥ B-tokens) → Chrome convergence (A4 ⊗ B-chrome) → A: collab/attach + thin host`.

---

## 3. Track A — SDK restructure plan

Target contract is [`SDK_ARCHITECTURE.md`](./SDK_ARCHITECTURE.md) lines 126-137.
Working rules from [`SDK_MIGRATION_PIPELINE.md`](./SDK_MIGRATION_PIPELINE.md) apply:
3–4 commits → green CI; no semver break without a changeset; UI changes
Playwright-gated.

### Phase 0 — Univer fork 0.24 → 0.25 *(blocking — gates BOTH tracks)*

This is the true gate. `packages/sdk/package.json` and `apps/web/package.json` pin
`0.25.0`, but the **actual install is 0.24.0** (`node_modules/.pnpm/@univerjs+*@0.24.0`,
vendor symlinks → `vendor/univer-revamp` on `casual-sheets/0.24`). The DI work in
Batch 2 must be validated on the version actually installed. Finish the
0.24→0.25 fork upgrade (see `SDK_MIGRATION_PIPELINE.md` Phase 0) before any
restructure or DS chrome work lands.

**Exit:** `vendor/univer-revamp` on `casual-sheets/0.25`, the six custom commits
cherry-picked, install resolves to 0.25.0, app builds and smoke-passes on 0.25.

### Batch 1 — lazy-plugins lift *(LOW risk — mostly done)*

**Done:** `lazy-plugins.ts` lives at `packages/sdk/src/univer/lazy-plugins.ts`,
re-exported via the `./univer` subpath (`packages/sdk/src/index.ts:19`,
`package.json` exports lines 46-50). App imports it from
`@casualoffice/sheets/univer` (`UniverSheet.tsx:24`, `dev-helpers.ts:5`). The
module-level singleton (`currentUniver`/`loaded`/`inflight`) is SDK-owned
(`index.ts:16-19`) so host and SDK resolve one instance.

**Remaining:** keep `lazy-plugins` `@internal` (not semver-covered) through the
editor move. Resolve the version skew — pins say 0.25.0, install is 0.24.0
(folded into Phase 0).

**Exit:** Phase 0 done; install == pins; `lazy-plugins` internal-only and shared.

### Batch 2 — formula + plugin registration *(HIGH risk: DI + worker bundling)*

This is the crux. **Root cause of the DI blocker:** `IRPCChannelService` is a DI
token from `@univerjs/rpc` (`vendor/univer-revamp/packages/rpc/src/services/rpc/channel.service.ts:27`),
registered **only** by `UniverRPCMainThreadPlugin.onStarting()`
(`vendor/univer-revamp/packages/rpc/src/plugin.ts:90-102`), which itself requires
a `workerURL`/`Worker` (plugin.ts:81-86, throws without one).
`UniverSheetsFormulaPlugin` resolves that token **conditionally — only when
`notExecuteFormula:true`** (`vendor/univer-revamp/packages/sheets-formula/src/plugin.ts:125-130`):

```ts
// sheets-formula/src/plugin.ts:125-130
if (this._config.notExecuteFormula) {
  const rpcChannelService = j.get(IRPCChannelService);   // ← throws here
  dependencies.push([IRemoteRegisterFunctionService, {
    useFactory: () => toModule(rpcChannelService.requestChannel(...)),
  }]);
}
```

The SDK is in the worst configuration: `CasualSheets.tsx:143` sets
`notExecuteFormula:true` (which *forces* the lookup) but never registers
`UniverRPCMainThreadPlugin` (no worker bundled). Result:
`[redi]: Expect 1 dependency item(s) for id IRPCChannelService but get 0` — which
is why both `UniverSheetsFormulaPlugin` and `UniverSheetsFormulaUIPlugin` are
commented out (`CasualSheets.tsx:145-153`) and the SDK can't compute. The app avoids
it because `apps/web/src/univer/plugins.ts:43` registers
`UniverRPCMainThreadPlugin` with a real worker *before* the formula plugin.

**Two fix paths, sequenced:**

- **(B) Main-thread compute — land first, smallest change.** Set
  `notExecuteFormula:false` on engine + sheets + formula plugins; drop
  `UniverRPCMainThreadPlugin` entirely. The `IRPCChannelService` lookup is gated
  behind `notExecuteFormula` (plugin.ts:125), so with `false` it is never touched
  and DI succeeds. Re-add formula-UI. Cost: compute on the main thread (the thing
  the app worker avoids for large workbooks). Milestone: **"SDK computes formulas."**

- **(A) Worker-backed compute — layer in second for parity.** Move
  `apps/web/src/univer/plugins.ts` + `formula-worker.ts` into the SDK; register
  `UniverRPCMainThreadPlugin` with the worker; keep `notExecuteFormula:true`.
  **Concrete prerequisite: add `@univerjs/rpc` to the SDK's peer/dev deps — it is
  currently absent from `packages/sdk/package.json`** (only `apps/web/package.json:25`
  declares it). This is the one missing SDK dependency for the worker path. Risk:
  bundling a Web Worker inside a published library —
  `new Worker(new URL('./formula-worker.ts', import.meta.url), {type:'module'})`
  (plugins.ts:39) must survive `tsup` and the consumer's bundler. Needs an explicit
  worker-build strategy (tsup worker entry, or ship the worker as a separate built
  asset and let the host pass `workerURL`).

**Also in this batch (LOW risk):**
- `apps/web/src/univer/facade.ts` → SDK (10 `/facade` side-effect imports). Deps
  (`sheets-table/facade`, etc.) are already SDK deps (package.json:66-104).
- `apps/web/src/univer/styles.ts` → fold into SDK `./styles` entry (SDK already has
  `src/styles.ts`; CSS list at styles.ts:8-25).

**Exit:** SDK computes formulas (path B green), `@univerjs/rpc` added, worker path
(A) green in a published-tarball test, facade + styles folded in.

### Batch 3 — editor core + imperative API *(MEDIUM-HIGH risk)*

Lift the `UniverSheet.tsx` body into the SDK editor core; the current minimal
`CasualSheets.tsx` boot becomes the floor of a `<CasualSheets chrome="none">`.

**Moves:**
- Snapshot swap chain (`UniverSheet.tsx:152-227`) → `loadSnapshot`. The
  `createUniverSheet`/`disposeUnit` collision guards (`UniverSheet.tsx:179-227`)
  become the swap implementation.
- `runInitialRecalc` (`UniverSheet.tsx:96,259`) — forces compute on templates with
  no cached `<v>`.
- `dev-helpers.ts`, `disable-zoom-shortcut.ts`, `paste-merge-hook.ts` → SDK. These
  reach through `(api as any)._injector` (`paste-merge-hook.ts:30`,
  `disable-zoom-shortcut.ts:28`, `dev-helpers.ts:73`) — the unstable escape hatch;
  works but couples to Univer internals (MEDIUM risk).
- `extendContextMenu` Merge/Unmerge (`UniverSheet.tsx:83`).

**`CasualSheetsAPI` contract** (`SDK_ARCHITECTURE.md:126-137`):
`getSnapshot / loadSnapshot / importXlsx / exportXlsx / getSelection /
executeCommand / setTheme / attachCollab / univer`. **`attachCollab` is a
forward-declaration** — implementation is Phase 2; ship it as a stub or omit until
the collab batch. This is the **semver surface** — changeset required
(pipeline rule line 19).

**React seam:** `use-univer.tsx` / `UniverContext` is consumed by **21 shell
files** via `useUniverAPI` (`use-univer.tsx:6`). It must keep working — either
re-export from the SDK or have the host wrap `onReady(api)` back into its existing
context (MEDIUM risk: broad blast radius).

**Exit:** `<CasualSheets chrome="none">` exposes the full `CasualSheetsAPI`; app
drives the editor through it; changeset filed; all 21 `useUniverAPI` consumers
green.

### Batch 4 — chrome extraction *(HIGH risk — this is the convergence pass, see §5)*

Deferred to §5 because it is shared with Track B.

---

## 4. Track B — design-system redesign plan

**Package:** `@schnsrw/design-system` `0.1.0` (MIT, ESM-only, zero runtime deps,
React 18 peer). Tokens at `services/design-system/src/tokens/`; primitives at
`services/design-system/src/components/`. It is a faithful extraction of the
sheet's `styles.css` — token adoption is a dedupe.

### B0 — Resolve the dependency model BEFORE any code (the blocker that killed both prior attempts)

This is non-negotiable and comes first. From §7: both reverted attempts foundered
because the package is **not published to npm**. Rules:

- **Publish `@schnsrw/design-system` (even a private `0.1.0`) to a registry first.**
  Do not adopt it as a normal dependency until it resolves on a real registry.
- **Never commit `link:` / `file:` / machine-relative paths** in
  `apps/web/package.json` or `packages/sdk/package.json`. The
  `vendor/univer-revamp` override trick does **not** generalize — that target is
  *inside* the repo; `services/design-system` is a sibling *outside* the workspace.
  `^0.1.0` + a local `pnpm.overrides → link:` is the exact trap that passes locally
  and fails on CI / clone / tarball.
- **Decide the split up front (don't patch toward it):**
  **embed-runtime = inlined tokens, no runtime design-system dep; `/shell`
  components = design-system as an optional `peerDependency`** so consumers bring
  their own version. This is what the two reverted fixes were groping toward
  (`4193d51`, `904f402`); bake it in from the start.
- **Verify with a packed tarball, not just in-repo build.** Add `pnpm pack` →
  install into a clean dir → Drive consumes it, as a gate. Both reverts passed
  in-repo typecheck/build/smoke and still shipped an empty CSS bundle and a broken
  tarball.

### B1 — Fix the `editor-theme.css` `:root` leak in the package

The single biggest *token* adoption blocker. `tokens/editor-theme.css:12` targets
`:root, [data-app='docs']` and is imported last (`tokens.css:15`), so its first
selector **unconditionally recolors accent to docs-cyan (#0891b2) for everyone** —
the sheet's teal brand would silently turn cyan. Fix in the package (scope to
`[data-app='docs']` only) before the sheet imports `tokens.css`, or re-assert
`--color-accent` after import. Fix-in-package is correct.

### B2 — Token dedupe (parallel-safe with Track A)

The sheet already has identical tokens inline (`styles.css:5-160`:
`--color-accent: #0e7490`, `--titlebar-h: 64px`, `--toolbar-h: 66px`). Import
`@schnsrw/design-system/tokens.css` once at app entry, then **delete the now-duplicated
token block** from `styles.css`, keeping only sheet-specific tokens not in the
package. Caveat: package fonts load from Google Fonts CDN at runtime
(`tokens/fonts.css:15`) — for offline/Docker the sheet may want self-hosted fonts;
decide before deleting the app's own `index.html` font links.

### B3 — Presentational primitive restyle (parallel-safe; least-coupled first)

Restyle the pure-presentational primitives in `apps/web` to the package, in
analysis-3's least→most coupled order. These read no Univer plumbing and no
shell-coupling, so they're safe to run alongside Track A infra:

1. **`Icon.tsx`** (no app context) → back with package `Icon` (Material Symbols
   ligatures; matches the "Material Symbols Outlined only" rule). Color already
   inherits via `currentColor`.
2. **`Dialog.tsx` / `Popover.tsx` / `Tooltip.tsx`** — restyle the shared modal/
   popover/tooltip frames; all dialogs + dropdowns inherit. **Caveat:** the
   package's `Dialog`/`Menu`/`Tooltip` are *less capable* than the sheet's (no
   ESC/focus-trap/portal/collision/flip). **Adopt the DS overlays only as styled
   panels behind the app's existing positioning/focus logic — not as wholesale
   replacements.** Keep the sheet's `Popover.tsx` as the anchor.
3. **`RibbonControls.tsx`** — the single best leverage point: imports only
   `Icon`/`Popover`/`Tooltip`, everything via props
   (`RibbonControls.tsx:1-4`). Restyling its `RibbonGroup`/`BigToolbarButton`/
   `ToolbarButton`/`ToolbarSelect`/`ToolbarDropdown` (via DS `IconButton`/`Select`/
   `Button`) re-skins the entire Toolbar without editing `Toolbar.tsx`.
4. **Status pills** (`BusyPill`/`SaveStatusPill`/`ActivityPill`/`NamePill`) — one
   context read each; back with DS `Pill`/`Badge`. **Close the inline-hex leaks**
   here (`#0f172a`, `#e2e8f0`, `#d93025`, `#188038` in `ActivityPill.tsx`,
   `HistoryPanel.tsx`, `SaveStatusPill.tsx`, `NamePill.tsx`, etc.) — they don't
   re-theme in dark mode.

**Stop here for the parallel phase.** `FormulaBar`, `SheetTabs`/status bar,
`MobileActionBar`, `TitleBar`, `MenuBar` are *not* restyled in place — they go
through the convergence pass (§5) to avoid double-touch.

### Light/dark theming

Light = `:root` default; dark = manual `[data-theme='dark']` opt-in (no
`prefers-color-scheme`). The sheet already bridges this to the Univer canvas via a
body class (`ThemeBridge.tsx` / `theme.ts`, toggled from `TitleBar.tsx:187`).
Per-product accent via `data-app="docs"`. Components reference only tokens, so a
theme/accent swap propagates without JS re-render. **One discipline gap:** DS focus
rings fire on `:focus` (mouse + keyboard) via JS `useState`, not `:focus-visible` —
mouse clicks flash the glow. Wrap or override to `:focus-visible` for any DS
primitive that lands in the chrome.

**Component → chrome mapping** (drives §5): `Button`→dialog footers + share CTA;
`IconButton`(`pressed`)→toolbar format toggles; `Select`→font/size/number-format
dropdowns; `Tabs`→ribbon Home/Insert/Data strip + panel switchers;
`Pill`→Save/Activity/Busy/Name pills + toolbar range pill;
`Avatar`/`AvatarStack`(`active` green ring)→`CollabIndicator` presence;
`Card`→home file-picker tiles; `Kbd`→`KeyboardShortcutsDialog` + palette chips.
The **chrome-height tokens** (`--titlebar-h`/`--toolbar-h`/`--formula-bar-h`/
`--sheet-tabs-h`/`--statusbar-h`, `tokens/spacing.css:57-61`) are the embedding
contract — they reproduce the sheet's exact `grid-template-rows`
(`styles.css:322-323, 345-346`) so sheet and docs render identical strip geometry.

---

## 5. The convergence — chrome into the SDK as design-system slot components

**This is Track A Batch 4 and Track B's chrome restyle, done as one pass.** Each
shell file is moved into `packages/sdk/src/chrome/*` *and* re-skinned with DS
primitives *and* has its app-coupling inverted to props/slots — in a single edit.
HIGH risk, Playwright-gated. The reverted attempt 1 (§7) already ported
`TitleBar`/`Toolbar`/`FormulaBar`/`SheetTabs`/`StatusBar`/`SheetShell` (~1179 LOC)
to `packages/sdk/src/shell/` as props-driven `forwardRef` components — that shape
is the model; it was dragged down only by the dependency model, which B0 fixes.

### Cut the four import seams first (chrome cannot hard-import host concerns)

Before any file moves, invert these seams so the chrome receives them as
props/slots, not imports:

| Seam | Files | Inversion |
|---|---|---|
| **API access** | 21 shell files via `useUniverAPI` (`use-univer.tsx:6`) | SDK-owned context or the `CasualSheetsAPI` ref (forward from Batch 3). |
| **Collab** (×7) | `TitleBar.tsx:5,6`, `MenuBar.tsx:37`, `CollabIndicator.tsx:2`, `NamePill.tsx:2`, `PreviewDriver.tsx:4`, `HistoryPanel.tsx:3,4`, `local-history.ts:5` | Collab is Phase-2 opt-in/attached; chrome takes presence/collab state as props, never imports `../collab/*`. |
| **FileSource** | `file-actions.ts:21,22,:447` (`selectFileSource().save`) | Save/open become host-injected callbacks, not a `../file-source` import. |
| **Auth** | `TitleBar.tsx:14` (`AccountMenu`) | `AccountMenu` becomes a slot; host injects it. |
| **Routing** | `TitleBar.tsx:15` (`navigate`) | Navigation becomes a callback prop. |

`save-status-context.tsx:60` and `activity-context.tsx:60` are chrome-internal
contexts — they move with the chrome cleanly.

### Order: lowest-coupling chrome first, highest last

Move-and-reskin in this order (mirrors analysis-3's coupling ladder and analysis-1
§5):

1. **`FormulaBar.tsx`** — only `useUniverAPI` + `useActiveCellState`; DS `Input` for
   formula box + Name Box.
2. **Dialogs** (`FormatCellsDialog`, `GoalSeekDialog`, `PageSetupDialog`,
   `InsertCellsDialog`, `PasteSpecialDialog`, `NameManagerDialog`,
   `PropertiesDialog`, `AboutDialog`, etc.) — DS `Button` footers, `Input`/`Select`/
   `Checkbox`/`Switch` fields, `Tabs` for category tabs, behind the app's
   `Dialog.tsx` frame (keep ESC/focus-trap).
3. **`SheetTabs.tsx` + status bar** (status bar is rendered inside `SheetTabs.tsx`,
   zoom at `:266-305`) + **`MobileActionBar.tsx`** — DS `Pill`/`IconButton`.
4. **`Toolbar.tsx`** — visuals already delegated to `RibbonControls` (restyled in
   B3), so this is mostly the move + ribbon `Tabs`.
5. **`CollabIndicator.tsx`** — DS `Avatar`/`AvatarStack` (`active` green ring);
   collab state via props.
6. **`TitleBar.tsx`** — highest coupling (collab + auth + router + theme +
   workbook). Move last; `AccountMenu` and `navigate` as slots/callbacks.
7. **`MenuBar.tsx`** — the deepest hub (~75 imports, 2427 lines). Move last; DS
   `Menu` as styled panels behind the app's positioning. Treat as integration work,
   not pure CSS.

Expose behind `<CasualSheets chrome="full"|"minimal"|"none">`. `chrome="none"` is
the Batch-3 floor; `full` composes the slot stack
(`TitleBar → Toolbar → banner → FormulaBar → grid-row → MobileActionBar →
SheetTabs`, per `App.tsx:436-464`).

**Exit:** `shell/*` lives in `packages/sdk/src/chrome/*` as DS-skinned slot
components; `apps/web` injects FileSource/collab/auth/routing via props/slots;
Playwright green for every moved surface; changeset filed (chrome is semver
surface).

---

## 6. Combined milestones + ordering

One ordered list. `∥` = parallel-safe (separate branches); `→` = serial.

1. **`[serial]` Phase 0 — Univer fork 0.24 → 0.25.** Blocks everything. Install
   must equal the 0.25 pins. *(Track A)*
2. **`[serial]` B0 — publish `@schnsrw/design-system` to a registry** + decide the
   embed-runtime(inlined)/`/shell`(optional peer) split + `pnpm pack` consumer gate.
   Blocks all DS code. *(Track B)*
3. **`[serial]` B1 — fix `editor-theme.css` `:root` accent leak in the package.**
   *(Track B)*
4. **`[∥ A]` Batch 1 finalize** (lazy-plugins internal, install==pins) **and Batch 2
   path B** (`notExecuteFormula:false`, "SDK computes formulas"). *(Track A)*
5. **`[∥ B]` B2 token dedupe** + **B3 presentational restyle** (`Icon` → overlays →
   `RibbonControls` → status pills, inline-hex cleanup). *(Track B)* — runs
   alongside step 4; disjoint files.
6. **`[serial after 4]` Batch 2 path A — worker compute parity.** Add
   `@univerjs/rpc` to SDK deps; worker build via tsup; tarball test. *(Track A)*
7. **`[serial after 4,6]` Batch 3 — editor core + `CasualSheetsAPI`** (swap chain,
   recalc, dev-helpers, `_injector` helpers; `attachCollab` stub; keep the 21
   `useUniverAPI` consumers green; changeset). *(Track A)*
8. **`[serial — CONVERGENCE, needs 3,5,7]` Chrome pass (§5).** Cut the four seams;
   move-and-reskin `FormulaBar → dialogs → SheetTabs/status/MobileActionBar →
   Toolbar → CollabIndicator → TitleBar → MenuBar` into `packages/sdk/src/chrome/*`
   as DS slot components. Playwright-gated per surface. *(A4 ⊗ B-chrome)*
9. **`[serial]` Collab `attachCollab` implementation** (Phase 2) — fill the
   forward-declared API member; wire collab state into the now-slotted chrome.
10. **`[serial]` Slim `apps/web` to a thin host** consuming `@casualoffice/sheets`
    `<CasualSheets chrome="full">` with injected FileSource/auth/routing.

**Parallel windows:** steps 4 ∥ 5 is the one true parallel window (SDK infra vs
design tokens). Everything from step 8 onward is serial — the chrome is the funnel.

---

## 7. Risks & how this avoids the previous revert

**What happened before** (analysis 4): two adoption attempts, both reverted within
minutes. Attempt 0 (`bc42d9b`, reverted `9489c37` 12 min later) inlined design
values into `styles.css`. Attempt 1 (5-commit stack, all reverted at `c6bee29`…
`ed5a0db` 13 min after the last fix) wired the real package and hit, in order:

1. **tsup couldn't bundle bare-specifier CSS** — `import '@schnsrw/design-system/tokens.css'`
   produced an embed-runtime with **zero token references**, silently (iframe
   renders unstyled). Fix `4193d51` inlined tokens as critical CSS.
2. **The published tarball carried an unresolvable `link:` path** — `link:../../../design-system`
   baked into the packed `@casualoffice/sheets` tarball pointed at the author's
   machine; Drive's `file:` install broke. Fix `904f402` moved to optional dep +
   override.
3. **Machine-specific paths committed across 3 manifests** broke CI / fresh clone.
4. **Root cause both fixes orbited: the package was never published.** No dependency
   expression worked simultaneously for local dev, CI/clone, *and* the published
   tarball. Even the fixes were reverted — the fault was structural, not a bug.

**How this plan avoids it:**

| Prior failure | This plan |
|---|---|
| Unpublished package as a hard dep | **B0 first:** publish to a registry before any code; no adoption until it resolves there. |
| `link:`/`file:`/machine paths committed | **Banned.** The `vendor/univer-revamp` override trick explicitly does not generalize to an out-of-workspace sibling. |
| tsup silently drops bare-specifier CSS | **Embed-runtime = inlined tokens, no runtime DS dep** — decided up front, not patched. Add a build-from-package + CI-divergence check to prevent token drift. |
| Tarball broke at the consumer | **`pnpm pack` → clean-dir install → Drive-consumes** is a gate before "done." In-repo green is insufficient. |
| Shipped a changeset that wasn't shippable | **No version-bump while the dep is unresolvable** (the prior `13ce630` correctly refused). |
| `/shell` dragged down by the dep model | **`/shell` = optional `peerDependency`**; consumers bring their own DS version. The shell components can land independently once that's true. |

**New risks introduced by this plan:**
- **Worker-in-a-published-library (Batch 2A).** `new Worker(new URL(...))` may not
  survive tsup + the consumer bundler. Mitigation: ship path B (main-thread) first;
  treat worker as a perf follow-up behind a tarball test.
- **`_injector` escape hatch** (`paste-merge-hook.ts:30` et al.) couples the SDK to
  Univer internals across the 0.25 line. Mitigation: validate on the 0.25 fork in
  Batch 3; the fork is ours to patch if the internal shifts.
- **21-file `useUniverAPI` blast radius** in Batch 3. Mitigation: re-export the
  context from the SDK so consumers don't change import paths in lockstep.
- **Convergence-pass size.** ~70 shell files, Playwright-gated. Mitigation: the
  lowest→highest coupling order (§5) keeps each commit batch to 3–4 surfaces;
  `MenuBar`/`TitleBar` land alone.

---

## 8. Multi-agent execution note

Good candidates for parallel agent fan-out:

- **Step 4 ∥ Step 5 — the headline parallel window.** One agent on Track A
  Batch 2 path B (SDK formula DI, `packages/sdk/src/univer/*`,
  `CasualSheets.tsx`); a second agent on Track B B2/B3 (token dedupe +
  presentational restyle in `apps/web/src/styles.css` + `Icon`/overlays/
  `RibbonControls`/pills). Disjoint file sets — no merge conflict.
- **Within B3 — primitive restyles fan out.** `Icon`, the overlay frames
  (`Dialog`/`Popover`/`Tooltip`), `RibbonControls`, and each status pill are
  independent files; one agent each, converge on review. Caveat: keep the overlay
  agent from replacing positioning/focus logic (DS overlays are panels only).
- **Within step 8 (convergence) — the early, low-coupling surfaces fan out** once
  the four seams are cut: `FormulaBar`, the dialog set, `SheetTabs`/status bar,
  `MobileActionBar` are largely independent. **Do not** fan out `TitleBar` and
  `MenuBar` — they are the integration nexus (collab + auth + router + ~75 imports);
  one agent, serial, last.

**Keep serial (single agent):** Phase 0 (fork upgrade), B0 (dependency model +
publish), Batch 3 (the semver `CasualSheetsAPI` surface + 21-consumer seam), the
seam-cutting prelude to step 8, and `attachCollab` (step 9). These define
contracts the parallel work depends on; splitting them invites contract drift.
