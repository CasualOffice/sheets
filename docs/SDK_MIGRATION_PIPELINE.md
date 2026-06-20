# SDK Migration Pipeline

The staged path from today's structure to the [`SDK_ARCHITECTURE.md`](./SDK_ARCHITECTURE.md)
target: full editor in the SDK, opt-in storage/collab adapters, a thin `apps/web` host.

**Direction locked (2026-06-19):** full editor into the SDK · slim `apps/web` onto it ·
Univer **0.25 first** (Phase 0).

## Working rules (apply to every phase)

- **Small batches → CI checkpoint.** 3–4 commits, push, wait for green CI before piling
  on more. Never stack a second batch on a red one.
- **UI changes are gated on Playwright + CI.** Any change that touches rendered UI must
  be driven through Playwright (observe the real screen/flow) **and** pass CI before it
  reaches origin. Typecheck + unit alone is not sufficient for UI.
- **Full local validation before push:** `lint` + `format:check` + `typecheck` +
  `test:unit` + `build` (+ relevant `test:e2e`). See `SKILLS.md`.
- **No semver break without a changeset.** The SDK is published; props +
  `CasualSheetsAPI` are the contract.

---

## Phase 0 — Univer fork 0.24 → 0.25  *(✅ done)*

**Done.** The vendored submodule `vendor/univer-revamp` (remote `CasualOffice/univer-revamp`)
now sits on `casual-sheets/0.25`, at the `v0.25.0` release with the **six** custom commits
cherry-picked on top. Every `@univerjs/*` pin is `0.25.0` (apps/web + packages/sdk) and the
`pnpm.overrides` block links to the fork. The steps below are kept as the recipe for the
*next* upgrade (e.g. 0.26).

The six custom commits:
- 2 feature: paste-merge formula/row-property preservation; filtered-dropdown visibility.
- 4 perf (see `UNIVER_FORK_PERF.md`): font-measure cache LRU; merge-range row-bucket
  index; header hit-test row/column index; setStylesCache visible-span walk.

**Steps**
1. In `vendor/univer-revamp`: add `upstream` (dream-num), fetch `v0.25.0`, create
   `casual-sheets/0.25` from the tag.
2. Cherry-pick the six custom commits (`0bcc094b8..52d85ec78`) onto the new branch;
   resolve conflicts against 0.25 APIs.
3. Bump every `@univerjs/*` pin `0.24.0 → 0.25.0` in `apps/web/package.json` and
   `packages/sdk/package.json` (peer + dev). Keep them **exact and identical** — never
   mix versions.
4. Re-audit the `pnpm.overrides` block in root `package.json`: every `@univerjs/*` we
   import (direct + transitive) must be listed, or it silently resolves to npm 0.24 and
   DI breaks with `Service2`-style suffix errors. Check the 0.25 transitive tree.
5. `pnpm install`, rebuild, run the full suite.

**Exit criteria:** `IWorkbookData` round-trips unchanged; collab bridge mutations
replay; xlsx import/export byte-stable on the fixture set; full Playwright suite green;
no duplicate-redi runtime errors.

**Milestone M0:** fork on `casual-sheets/0.25`, app + SDK pinned 0.25.0, CI green.

---

## Phase 0.5 — Wire missing Univer integrations  *(feature parity from the lib we vendor)*

The fork ships capabilities we never integrated. Add them *after* the 0.25 base is
stable and *before* the SDK restructure — each is a drop-in `Univer*Plugin`.

| Plugin | Capability | Priority |
| --- | --- | --- |
| `UniverSheetsCrosshairHighlightPlugin` | Excel active row/column crosshair highlight | high (Excel parity) |
| `UniverSheetsZenEditorPlugin` | Immersive / full-screen cell editor for long content | high |
| `UniverSheetsGraphicsPlugin` | In-cell graphics — evaluate vs our custom `apps/web/src/sparklines/` | medium |
| `UniverWatermarkPlugin` | Confidential watermark overlay (personal/team/WOPI modes) | medium |
| `UniverActionRecorderPlugin` | Record & replay user actions (macro-style) | low |

Out of scope by policy: `uniscript` (scripting — deferred), `docs-*`/`slides*` (other
editors), Vue/web-component adapters. The doc editor is **not** Univer-based.

**Exit criteria:** each wired plugin verified via Playwright (Excel-parity behavior),
CI green. Crosshair + zen-editor first.

**Status (2026-06-19):** ✅ crosshair-highlight + zen-editor wired (eager, context-menu
triggers) — `tests/e2e/univer-extras.spec.ts`. ⏸️ graphics/watermark/action-recorder
**deferred**: not drop-ins. `sheets-graphics` is a render primitive with no standalone
UI (overlaps our custom sparklines); `watermark` + `action-recorder` target Univer's
ribbon (which we hide) and need custom-shell triggers + product UX — each becomes its own
scoped feature PR.

**Milestone M0.5:** ✅ crosshair-highlight + zen-editor shipped; the other three scoped
out to dedicated PRs.

## Phase 1 — Promote the full editor into the SDK  *(G1, G2)*

Move the real editor out of `apps/web` and into `@casualoffice/sheets` so the package
*is* the product.

**Steps**
1. Lift `apps/web/src/UniverSheet.tsx` (all lazy plugins, paste/merge hooks, formula
   worker, snapshot swap) into `packages/sdk/src/sheets/` as the new `<CasualSheets>`
   core. The old minimal boot becomes `chrome="none"`.
2. Lift the Office chrome (`apps/web/src/shell/`: Ribbon, FormulaBar, StatusBar,
   TitleBar, FileMenu) into `@casualoffice/sheets/chrome` as **slot components** behind
   the `chrome="full" | "minimal" | "none"` prop.
3. Define and implement the `CasualSheetsAPI` imperative ref + the props contract from
   `SDK_ARCHITECTURE.md` (`getSnapshot`/`loadSnapshot`/`importXlsx`/`exportXlsx`/
   `executeCommand`/`setTheme`/`attachCollab`/`univer`).
4. Keep `apps/web` compiling by re-importing from the SDK at each step (no big-bang).
5. Add a changeset (minor bump — additive API).

**Exit criteria:** the SDK alone renders the full editor in the embed playground;
`apps/web` consumes only SDK exports for the editor + chrome; props/API documented;
Playwright embed spec green.

**Milestone M1:** `@casualoffice/sheets` renders the complete Office editor with a
documented props + `CasualSheetsAPI` surface.

---

## Phase 2 — Save/exit event contract + opt-in collab  *(G3)*

The SDK persists **nothing**. This phase formalizes the **host-owned persistence
contract** (the editor hands data out on change/save/exit; the host stores it) and
makes collab opt-in with **WOPI-backed** persistence. There is **no
`BrowserFileSource` / localStorage built into the SDK** — that earlier framing was
wrong (localStorage is a *demo-host* choice, not an SDK feature; see
`SDK_ARCHITECTURE.md` › Storage).

**Steps**
1. **Define the save/exit event contract** on the SDK, delivered two ways with one
   shape: React hooks (`onChange` / `onSave` / `onExit`) and **postMessage** for the
   `<iframe>` embed (extend `embed-runtime`'s `EmbedTransport` with `save`/`exit`).
   The SDK never writes a store; it only emits the snapshot.
2. **Persistence stays host-side.** Today's `apps/web/src/file-source/` (WOPI,
   personal, demo localStorage) becomes a *host* consumer of those events, not an
   SDK-bundled `FileSource`. The demo (`apps/web` on Pages) wires events →
   localStorage; real hosts wire events → WOPI / their backend.
3. Move `apps/web/src/collab/` (bridge, presence, driver) to
   `@casualoffice/sheets/collab`. Expose `attachCollab(api, { room, server, password })`
   returning a detach handle. The editor stays collab-unaware until attached.
   **In collab mode the authoritative document is saved via WOPI / host protocol**,
   not a browser store; Yjs/Hocuspocus is only the realtime transport.
4. Confirm the non-negotiable collab hooks survive the move: subscribe to
   `ICommandService.onMutationExecutedForCollab`; apply remote with
   `IExecutionOptions.fromCollab`; respect `params.__splitChunk__`.
5. `apps/server` unchanged.

**Exit criteria:** a host receives save/exit events (React + postMessage) and can
persist with no SDK-side storage; the Pages demo round-trips via localStorage as a
host; `attachCollab` opens a room against `apps/server` with WOPI-backed save;
WOPI + personal Playwright configs still green.

**Milestone M2:** save/exit event contract shipped on both delivery surfaces;
collab is one call with WOPI-backed persistence; the SDK stores nothing.

---

## Phase 3 — Slim `apps/web` into a thin reference host  *(G4)*

Make `apps/web` the excalidraw.com-equivalent: mostly SDK. As the backendless Pages
demo it persists to localStorage **as the host** (consuming the SDK's save/exit
events); WOPI/personal/collab are host-side layers for real deployments.

**Steps**
1. Reduce `apps/web` to: route shell + host persistence (save/exit events →
   localStorage for the demo, WOPI/personal otherwise) + mounting
   `<CasualSheets chrome="full">` + opt-in `attachCollab`.
2. Delete app-local copies now living in the SDK; import from the package.
3. Keep `/home` picker, path router, admin panel (personal mode) as host concerns — they
   are *host* features, not editor features.

**Exit criteria:** `apps/web` contains no editor/chrome/storage/collab logic of its own,
only composition; all existing app Playwright suites green; bundle is SDK-dominated.

**Milestone M3:** `apps/web` is a thin SDK consumer that doubles as the live integration
example.

---

## Phase 4 — Adopt `@schnsrw/design-system`  *(G4, shared look)*

The design system (`@schnsrw/design-system`, Inter + Material Symbols + tokens) exists
but sheet doesn't consume it.

**Steps**
1. Add the dependency; import `@schnsrw/design-system/tokens.css` once at the host root.
2. Migrate SDK chrome primitives (Button, IconButton, Dialog, Menu, Input, Badge, Pill,
   Avatar/Stack, Tooltip, Kbd) to the design-system components, replacing inline
   equivalents. Material Symbols (Outlined) only — never text glyphs.
3. Wire `setTheme` to the design system's `data-theme="dark"` toggle.
4. **Heavy Playwright pass** — this is the largest visual surface change. Screenshot the
   ribbon, formula bar, dialogs, menus, status bar in light + dark before/after.

**Exit criteria:** chrome renders entirely via design-system tokens/components; dark mode
works; visual Playwright review signed off; CI green.

**Milestone M4:** suite-consistent look; sheet shares the design system with doc/slides/
drive.

---

## Phase 5 — Integration docs + release

**Steps**
1. Write the integration guide: install, `<CasualSheets>` props, `CasualSheetsAPI`,
   storage adapters, `attachCollab`, the iframe path (`CasualSheetsIframe` /
   `embed-runtime`), and server-side `xlsx` usage.
2. Refresh `ARCHITECTURE.md` to describe the post-migration runtime (it predates
   `packages/sdk`).
3. Publish `@casualoffice/sheets` (and any graduated storage/collab packages) via the
   changeset release flow.
4. Add a copy-paste integration example app to `site/` demos.

**Milestone M5:** versioned SDK on npm with a complete integration guide; the suite site
demonstrates the embed.

---

## Milestone summary

| ID | Milestone | Gap closed |
| --- | --- | --- |
| M0 | Fork on `casual-sheets/0.25`, app+SDK pinned 0.25, CI green | — |
| M0.5 | Missing Univer plugins wired (crosshair, zen-editor first) | parity |
| M1 | SDK renders the full Office editor; documented props + `CasualSheetsAPI` | G1, G2 |
| M2 | Save/exit event contract (hooks + postMessage); host owns persistence; collab one call, WOPI-backed | G3 |
| M3 | `apps/web` is a thin SDK consumer | G4 |
| M4 | Design-system adopted; suite-consistent look + dark mode | G4 |
| M5 | Published SDK + integration guide + live embed example | — |

## Sequencing notes

- M0 is **blocking** — everything else builds on 0.25.
- M1 → M2 → M3 are ordered; M3 can only delete app-local code once M1/M2 export it.
- M4 (design system) can overlap M3 but lands its own heavy Playwright pass.
- M5 closes the loop; partial integration docs can ship earlier as the API stabilizes.
