# UX audit — Casual Sheets + Casual Editor (Phase C + Phase D state, 2026-06-11)

Honest audit, no excuses. Triggered by the user's report on the v0.3.1
single-mode flow ("no head, no tail"). The user noticed the visible
symptoms in 5 minutes; what follows is the structural root cause
and every related gap I found in a deeper pass through both repos.

Format: each issue carries **What's broken**, **Why (root cause)**,
**Industry standard** (Google Docs, Sheets, Office 365, Dropbox,
Notion — whatever's the closest peer), and **Fix shape** with a rough
size. Priorities at the bottom.

---

## §1 — Architectural root cause

Both repos boot **editor-first**. The editor is always mounted; the
"home" / "file list" / "auth gate" / "share landing" surfaces are
**modal overlays or dismissable banners on top of the editor**, not
separate routes.

Concrete evidence (sheet):

- `apps/web/src/main.tsx` routes only on `/admin/*` vs everything-else.
  Everything-else mounts `<App>`, which always mounts `<UniverRoot>` +
  `<UniverSheet>` with an `emptyWorkbook` snapshot the moment auth
  resolves.
- `<HomeScreen>` is rendered as a **sibling of the editor** with
  `dismissed` state (`apps/web/src/App.tsx:401`). It's an overlay, not
  a route.
- `<PersonalAuthGate>` wraps everything *including the always-mounted
  editor*. The user sees the auth gate, fills it, and the gate
  dismisses to reveal an editor with a blank workbook — not a file
  list.

Same shape in document (docx-editor):

- `examples/vite/src/App.tsx` toggles between `<Home>` and `<DocxEditor>`
  via local React state, not URL. No browser-back to file list.
- The deployed editor app makes the same trade — open URL ⇒ editor.

**Industry standard:** Google Docs, Sheets, Notion, Office.com, Dropbox
Paper all use route-distinct surfaces:

- `app.example.com/` ⇒ file list ("My Drive", "All sheets", "Workspaces")
- `app.example.com/d/<id>` (or `/sheets/<id>`) ⇒ editor for that document
- `app.example.com/share/<token>` ⇒ recipient-specific landing

The URL is the canonical place. Back/forward, copy-link, deep-link
all work. Auth gate doesn't render the editor in the background.

**Fix shape:** introduce a router (sheet: hash-based or
`history.pushState` — no react-router needed for v1; the existing
`/admin` pathname switch in `main.tsx` is the model). Three routes
minimum: `/`, `/r/<roomId>` (collab room, exists today), `/d/<fileId>`
(personal-mode file). Plus `/share/<token>` for share-link landing.
~3 files. **This is the prerequisite for everything else below.**

---

## §2 — Per-flow audit (sheet)

### §2.1 First-time sign-in (personal mode, single)

| | |
|---|---|
| **Today** | Visit → `<PersonalAuthGate>` shows a signup card → user signs up → gate dismisses → user lands in a **blank workbook**. No "Welcome", no file list, no orientation, no idea what to do. The first thing they see is an empty Univer grid with their cursor in A1. |
| **Root cause** | Editor mounts during auth resolution; gate dismissal reveals an already-running editor. There's no concept of "first-time user has no files yet → show empty state of My Files". |
| **Industry std** | Google Sheets first-time: list view with "Create new" + a templates row + a starter sample workbook. User sees "Sheet1 - Untitled" in the list immediately even before they create anything; clicking it opens the editor. |
| **Fix** | Land on `/` ⇒ **My Files** list view. Empty state: "No files yet — create one" + Templates row + a "New blank spreadsheet" button. Editor mounts only when user navigates to `/d/<id>`. |

### §2.2 "No way to see my files"

| | |
|---|---|
| **Today** | There IS a HomeScreen with a Recent Files section (`apps/web/src/home/HomeScreen.tsx:343`). But it's an **overlay that only opens when the user explicitly dismisses the editor** — and there's no obvious button to dismiss. The chrome shows Title Bar → Toolbar → Sheet grid; no "Home", no "All files", no breadcrumb back to a list. |
| **Root cause** | HomeScreen was designed for the anonymous flow ("pick a template to jump in"). It was never re-designed for personal mode where users have actual files they need to manage. |
| **Industry std** | Persistent left rail or breadcrumb. Google Sheets: top-left waffle / Sheets-home button → list view, always one click. Office Online: app launcher in the top bar. |
| **Fix** | The "click the logo to go home" pattern the user described **is the right pattern** — make the logo click → `/` (My Files list) instead of the current "show home overlay" behaviour. ~1 file in `TitleBar`. |

### §2.3 Save creates duplicate rows of the "same" file

| | |
|---|---|
| **Today** | User starts a New blank workbook → edits → Ctrl+S. `serverFileId` is `null` (workbook was created from default, not opened from a file row). Personal-file-source's `save()` falls through `if (opts.existingId)` and into `POST /files` (create path). The response sets a new id, **but the workbook meta never gets updated with the new `serverFileId`** — so the NEXT Ctrl+S falls through the same path → another `POST /files` → another row. Recent Files shows N rows of "Untitled.xlsx", all the same logical workbook. |
| **Root cause** | The create-save path in `apps/web/src/shell/file-actions.ts → saveAsXlsx` doesn't call `workbook.updateServerEtag` + the equivalent `updateServerFileId`. There IS a `WorkbookContext.updateServerEtag` for the etag side; the file-id side just isn't wired. |
| **Industry std** | Google Docs / Sheets: first save mints an id and binds the URL; subsequent saves overwrite. Never N copies. |
| **Fix** | After the create-save resolves, write `result.serverFileId` back into `workbook.meta`. **Belt-and-braces:** dedupe Recent Files by latest-modified-per-filename so the visible bug is gone even if the bind logic misses an edge case. ~2 files. |

### §2.4 Save creates duplicates ALSO on the autosave path

| | |
|---|---|
| **Today** | `useAutosave` ticks every N seconds. If the create-save bind in §2.3 is broken, every autosave tick creates yet another row. The user reports "saves multiple times" — partly the explicit save, partly autosave amplifying it. |
| **Root cause** | Same as §2.3. |
| **Industry std** | Autosave silently overwrites. |
| **Fix** | Same as §2.3 — one fix, two symptoms gone. |

### §2.5 Recent Files list has no notion of "open this" — it's a save log

| | |
|---|---|
| **Today** | Recent Files shows server file rows. Each is a separate POST result; user has no obvious way to pick "the latest" vs an older save of the same file. No filename grouping, no "last modified" sort, no thumb. |
| **Root cause** | The list was built for the anonymous browser flow (IDB-backed snapshots, each genuinely distinct). Personal-mode server files weren't re-thought. |
| **Industry std** | "Recent files" sorts by last-opened or last-modified DESC, one row per file, thumb on the left, filename + modified-at on the right. Click → open in editor. |
| **Fix** | Recent Files becomes the **MyFiles list** at `/`. Single source of truth: server file rows, sorted by `modified_at DESC`, dedupe by id, render thumb + name + modified-at + (rename / delete / share / download) overflow menu. ~1 component. |

### §2.6 "Click the logo multiple times → templates appear"

| | |
|---|---|
| **Today** | Logo's click handler likely cycles overlay state. User has to dismiss-then-reopen-then-flip-tabs to land on Templates. There's no canonical "templates" route. |
| **Root cause** | Templates live INSIDE the HomeScreen overlay as a section. No URL, no anchor. |
| **Industry std** | Templates have their own route or are a prominent section on the file list root. Google Sheets: Template gallery at the top of the file list, plus `Files → New → From template`. |
| **Fix** | Templates as a top section on `/` (My Files). Optional `/templates` route for the full gallery. Logo always → `/`. ~1 file. |

### §2.7 No admin panel affordance

| | |
|---|---|
| **Today** | `/admin` route exists (`AdminApp`), but no UI in the editor chrome links to it. Admins discover it only if someone tells them or they type the URL. |
| **Root cause** | Admin was treated as ops-only at first; the bootstrap-user flow makes the first signup an admin, so they expect a way to manage. |
| **Industry std** | Office 365 admin: top-right app launcher. Google Workspace admin: admin.google.com (separate domain but a clear menu link from any Workspace app). |
| **Fix** | `<AccountMenu>` (top-right user pill in `apps/web/src/auth/AccountMenu.tsx`) gets a new entry "Admin panel" when `me.role === 'admin'`. Click → navigate to `/admin`. ~1 file. |

### §2.8 Share-for-editing asks the recipient to log in (single mode)

| | |
|---|---|
| **Today** | "Share for co-editing" mints a share URL → recipient opens it → `<PersonalAuthGate>` blocks them with a sign-in card. In single mode there is **no second account they can use**, so they're stuck. |
| **Root cause** | `<PersonalAuthGate>` runs around the entire app, including share-link / collab-room URLs. There's no concept of "this route is share-link, skip auth". |
| **Industry std** | Google Sheets share-link: recipient lands directly on the document with view (or comment / edit) permissions. Auth optional. Anonymous editing supported via the link token. |
| **Fix** | `<PersonalAuthGate>` exempts `/r/<roomId>` and `/share/<token>` routes. In **single mode** specifically, sharing for **edit** is conceptually wrong — there's no "other user" who can hold the lock. Single-mode share should default to **view-only / download**. The "share for co-editing" entry in the menu either: (a) hides in single mode, (b) is enabled but flips to multi-mode personal accounts when the operator turns it on. ~2 files + a copy decision. |

### §2.9 Share-for-editing in multi mode — no recipient permission story

| | |
|---|---|
| **Today** | Multi mode supports per-user accounts. Sharing for editing today opens a room URL anyone with the link can join + edit. There's no per-recipient permission model (view vs comment vs edit), no recipient identity prompt, no audit of who edited what. |
| **Root cause** | Phase C personal accounts + the existing collab-room sharing flow weren't designed against each other. Rooms are anonymous; accounts are per-user. The intersection was left for later. |
| **Industry std** | Google Sheets: per-recipient or anyone-with-link, plus role (view / comment / edit), plus per-cell suggestion mode. Microsoft 365: identical. |
| **Fix** | Out-of-scope for a UX patch — needs a real design pass. Recommend a `docs/SHARING_MODEL.md` design before code. |

### §2.10 No file rename / move / delete from the file list

| | |
|---|---|
| **Today** | The editor has a "Rename" entry under File menu; there's no list view, so no list-level rename. Delete = browser back? No, just leave the editor. Tiny right-click context menu on Recent Files (HomeScreen) — but it's an overlay, not durable. |
| **Root cause** | No first-class file list. Same root cause as §2.5. |
| **Industry std** | File list rows: hover-show kebab → Rename / Make a copy / Move to / Move to bin / Download. |
| **Fix** | When My Files becomes a real route (§2.5), each row gets a kebab menu. ~1 component. |

### §2.11 First-save dialog confusion

| | |
|---|---|
| **Today** | New blank → Ctrl+S → no "Save As" dialog, no filename prompt, no folder picker. The workbook saves under the default name `workbook.xlsx`. User has no chance to name it before the row appears in Recent Files. |
| **Root cause** | `saveAsXlsx` defaults filename to `workbook.xlsx` (`apps/web/src/shell/file-actions.ts:174`). |
| **Industry std** | Google Sheets: title is editable in the title bar; clicking commits, Ctrl+S is silent (it's already saved). First save uses the current title. |
| **Fix** | Two halves: (a) make the title bar's filename **editable in place** (click → contenteditable → blur saves). (b) First-save uses the title-bar name; if it's the default "Untitled" leave it as `Untitled.xlsx` with a discoverable "Rename" prompt. ~2 files. |

### §2.12 No browser-tab title for the open document

| | |
|---|---|
| **Today** | `<title>` is hardcoded to "Casual Sheets" in `index.html`. Doesn't change per workbook. |
| **Industry std** | "Untitled spreadsheet - Google Sheets". Tab title tracks the document name. |
| **Fix** | `useEffect` on workbook meta → `document.title = nameOrUntitled + " — Casual Sheets"`. ~1 hook. |

### §2.13 Settings / Preferences are scattered

| | |
|---|---|
| **Today** | User-level prefs (display name, timezone, locale, avatar) live in `<SettingsModal>` (`apps/web/src/auth/SettingsModal.tsx`). Density / theme / formula-bar visibility are in scattered toolbar toggles. Profile edit is a separate dialog (`ProfileSettingsDialog`). |
| **Root cause** | Personal mode bolted on a profile system; the editor's existing prefs weren't unified. |
| **Industry std** | Single "Settings" page or modal. Sections: Account, Display, Language & region, Privacy, Notifications. One menu entry. |
| **Fix** | Consolidate. Account menu → Settings → one modal with tabs. ~1 file. |

### §2.14 Logout has no confirmation, no "any unsaved work?"

| | |
|---|---|
| **Today** | Account menu → Logout → page hard-reloads, browser is back to the sign-in card. If the user had unsaved changes in the open workbook, they're gone. |
| **Root cause** | Logout calls `POST /auth/logout` and reloads without checking dirty state. |
| **Industry std** | If unsaved: "You have unsaved changes — save before signing out?" modal. Else: silent. |
| **Fix** | Use the same dirty-check `<BeforeUnloadListener>` already wires for the close-tab case. ~1 file. |

### §2.15 Mobile UX

| | |
|---|---|
| **Today** | App is desktop-first. The mobile breakpoint shrinks the toolbar (`<MobileActionBar>`) but the **same single-page editor-first IA** is on mobile too. On a phone, the user can't see their file list, can't navigate, can't share — they get a tiny Univer grid. |
| **Root cause** | Single page, no routes; the mobile chrome is just a narrow toolbar. |
| **Industry std** | Mobile web Sheets opens directly to the file list. Editor is full-screen + chrome collapses. |
| **Fix** | Once §1 routing is in: `/` on mobile renders a list with thumbs; `/d/<id>` is the editor with mobile chrome. ~2 files (responsive sheet for `/`). |

### §2.16 Cmd-S vs Cmd-K vs Cmd-P keyboard shortcut discoverability

| | |
|---|---|
| **Today** | Shortcuts exist but `?` doesn't open a help modal. F1 doesn't either. |
| **Industry std** | `?` or `Ctrl+/` opens "Keyboard shortcuts" modal. Office: F1 → help. |
| **Fix** | `<HelpModal>` triggered by `?` key. Already mostly built in Drive — port the pattern. ~1 file. |

---

## §3 — Per-flow audit (document)

The architectural root cause is the same. Specific deltas:

### §3.1 Same editor-first IA

`examples/vite/src/App.tsx` toggles `<Home>` vs `<DocxEditor>` via React
state (`templateOrDoc !== null`), not routes. Browser back doesn't
return to Home. Refresh on an open document re-mounts Home. Reload-while-editing
loses the doc unless autosave caught it.

**Fix:** same as §1 — route-distinct. Sheet's pattern reusable.

### §3.2 PersonalAuthGate wraps everything

Same overlay-on-top-of-editor pattern. Same fix.

### §3.3 No "My Documents" list

Same gap as §2.2 / §2.5. The editor's recent-files surface is an
IDB-backed grid for the anonymous flow; personal-mode server files
aren't first-class.

**Fix:** the file-list component would parallel the sheet one; same
list with `/d/<id>` route, kebab menu, sort, dedupe.

### §3.4 Templates gallery is a separate landing

Templates show in `<Home>`. No route, no deep link.

**Fix:** templates as a section on My Documents root; `/templates`
optional.

### §3.5 Sharing model is identical

Same gap as §2.8 / §2.9 — single mode share-for-edit is broken;
multi mode lacks permissions. Same fix path.

### §3.6 Save mechanics also create duplicates

`@schnsrw/docx-js-editor`'s autosave uses `fileSource.save(id, bytes)`.
The Phase C PersonalFileSource has the same "id never bound after
first save" risk if the consumer doesn't write the result back to
its state. Drive's `DriveFileSource.save` does — but the editor's own
example apps/web doesn't. Worth double-checking.

### §3.7 Signing UI

Signing pane is a floating sidebar. Inside the editor it competes with
the toolbar. Mobile = no chance.

### §3.8 AI features (Translate / Writing Assistant / Citations)

Several dialogs (`TranslateDialog`, `WritingAssistantSheet`, `CitationsDialog`,
`TranslateDocumentDialog`) exist as separate components. There's no
single "AI" menu — each is hung off a different toolbar button or
keyboard shortcut. Users don't discover them.

**Industry std:** Google Docs "Tools → Help me write". Microsoft 365
"Copilot" panel. One entry-point, contextual sub-features.

**Fix:** an AI hub menu. Defer to a design pass.

---

## §4 — Shared issues

### §4.1 Error states are toast-only

Bytes failed to load → toast. Save conflict → toast. Save failed → toast.
After 3 seconds the toast vanishes; user has no log, no retry button,
no way to re-try the failed action without redoing it.

**Industry std:** persistent error chip + "Retry" / "Show details" /
"Copy error" actions. Activity log / Trash with deleted items
preserved 30 days.

**Fix:** activity log on the file list + retry chips on failed
operations. ~2 files.

### §4.2 No keyboard navigation between surfaces

Tab order within the editor is reasonable. Tab order from the editor
to the title bar / toolbar / sidebar is not. Power users can't reach
Save without Cmd-S, can't reach Files without mouse.

**Industry std:** Cmd-O for open, Cmd-Shift-S for save-as, Alt+/ for
menu access. Google Docs: Alt+/ opens the "Tell me" command palette.

**Fix:** command palette (Cmd+K) — already shipped in Drive, port the
pattern.

### §4.3 No offline / weak-network state

Personal mode requires the server. If the server is slow / offline,
autosave silently retries; user has no visual signal.

**Industry std:** "Saving…" / "Saved 2 min ago" / "Offline — will save
when connected" chip near the title.

**Fix:** integrate `<AutosaveStatus>` (already exists in the docx SDK)
prominently in `<TitleBar>`. ~1 file.

### §4.4 No "your name" pre-fill for collab

`<NamePrompt>` asks for a display name on join. Even if you're signed in
as `alice` in personal mode, the prompt still asks. Friction.

**Fix:** when authenticated, default to `me.username`. Skip the prompt
unless first-time on this device. ~1 file.

### §4.5 Avatars

`<AvatarStack>` renders colored circles with initials. Personal mode
has a profile-avatar URL but it doesn't surface in collab presence.

**Fix:** thread `profile.avatarUrl` through `presence.ts`. ~1 file.

---

## §5 — Prioritized fix plan

### Phase 1 — Foundational IA (unblocks everything else)  ✅ shipped 2026-06-11

1. ✅ **Path router** (sheet `apps/web/src/router/`, doc
   `docx-editor/examples/vite/src/router/`). pushState, no
   react-router. Routes: `/home`, `/sheet/<id>` (+ `/document/<id>`),
   `/sheet/new` (draft), `/r/<roomId>`. Commit `f50cbf6` (sheet) /
   `f823b0d` (doc).
2. ✅ **MySpreadsheetsList at `/home`** (sheet). Server file rows
   via `fs.listRecent` + `fs.subscribeRecent`, dedup by name, latest
   first. Commits `51352c3` + `11de5b7`.
3. ✅ **Editor at `/sheet/<id>`** — `RouteWorkbookSync` opens via
   `fileSource.openRecent`, binds `serverFileId` + `serverEtag` into
   `WorkbookMeta`. The first save mints an id and `replaceState`
   rebinds the URL from `/sheet/new` → `/sheet/<id>`. Commits
   `fe7c18f` + `6897dfb`.
4. ✅ **Logo → `/home`**, single click. Commit `06bb5a0`.
5. ✅ **Empty-draft Save skip**. `EditTracker` + `hasUserEdited` flag
   short-circuits create-save when the user has never typed.
   Commit `b89e93b`.
6. ✅ **Auth-gated `/` → `/home` redirect** so non-personal deploys
   (Pages, e2e) keep `/` rendering the editor. Commit `e50d87c`.

### Phase 2 — Personal mode chrome  ✅ shipped 2026-06-11

5. ✅ **AccountMenu → Admin entry** for `user.isAdmin`. Commit `96d39ee`.
6. ✅ **PersonalAuthGate exempts `/r/<roomId>`** — anonymous coedit
   joiners aren't asked to claim an account on the host's box.
   Commit `96d39ee`.
7. ⚖️ **Single-mode share = view-only / download.** Deferred per
   user direction: "for anonymous writer it's user's fault as he can
   share write access with password — architectural choice."
8. ✅ **Title-bar filename editable.** Already existed via TitleBar
   click-to-rename input.
9. ✅ **`document.title` tracks workbook name** — keyed off
   `route.kind` so back / forward flip the title immediately.
   Commit `96d39ee`.
10. ✅ **Logout dirty-check.** AccountMenu reads workbook context;
    `window.confirm` fires when `meta.hasUserEdited === true`.
    Commit `96d39ee`.

### Phase 3 — Mobile + a11y

11. ✅ **Mobile responsive `/home`.** `useIsMobile()` matchMedia
    hook, tighter padding, smaller title, wrapping actions row,
    Delete affordance always-visible on touch (no hover).
    Commit `2abb2ec`. Also fixed a FileSource-swap race that could
    flash an empty list on sign-in (same commit).
12. ✅ **Keyboard shortcuts cheat sheet** + `formatShortcut` util.
    Opens via `Ctrl+/`, `?` (outside an input), or Help menu.
    Renders ⌘ on Mac, `Ctrl+` on Win/Linux, friendly key names
    (PgUp / Esc). Commit `5962406`. MenuBar's existing entries now
    render through the same util — pays down the
    `feedback-shortcut-strings` memory debt. Commit `9f92397`.
13. ⏳ **Settings consolidation.** SettingsModal already has
    Profile / Security / Preferences tabs; the only "scattered"
    surfaces remaining are intentional (theme toggle as a one-click
    titlebar button, per-dialog dialogs). Closing unless a specific
    scattered thing surfaces.

### Phase 4 — Shared infra

14. ✅ **ActivityPill — persistent error log.** Bell icon + unread
    badge in TitleBar; popover lists last 25 errors newest-first with
    Dismiss per entry + "Clear all". Toast → activity bridge via a
    window `cd:activity-error` event so ToastContext stays unaware of
    ActivityContext. v1 ships log + dismissal; per-entry retry
    handlers deferred (needs per-call-site action context absent at
    the toast layer). Commit `b99beec`.
15. ✅ **Command-palette chord** — `Ctrl+Shift+P` (VS Code / Linear)
    aliased to the existing `CommandSearchDialog` (already bound to
    Alt+Q from Office). Three entry points, one dialog. Ctrl+K
    intentionally left as Insert Link (Excel convention).
    Commit `7c84605`.
16. ✅ **SaveStatusPill in TitleBar.** New `SaveStatusProvider` +
    `<SaveStatusPill>` shows Saving… → Saved {N ago} → Save failed
    with a relative-time tick. EditTracker drops the pill back to
    idle on the next mutation so it doesn't keep claiming "Saved 1
    hour ago" while the user is mid-typing. Commit `ea2b3fc`.
17. ✅ **Collab display name pre-fill.** CollabDriver reads the
    signed-in user via `useCurrentUser()` and pre-fills the room
    identity from `user.username` (write-through to localStorage so
    sign-out → anon-join keeps the same identity). Commit `43ed996`.

### Phase 5 — Needs design pass (NOT code yet)

18. ✅ **Multi-mode sharing permissions model.** §2.9. Shipped as
    `docs/SHARING_MODEL.md` — design only, no code. Commit `ea679a4`.
19. **AI features hub.** §3.8. `docs/AI_HUB.md`. Lives in the doc
    repo (sheet has no AI surfaces); deferred to the doc repo's own
    docs sweep.

---

## §6 — What the user will actually feel after Phase 1 (the minimum that fixes "no head no tail")

- Sign in → land on **My Files**. List of their workbooks, sorted by
  modified date, with a "New blank spreadsheet" button.
- Click a row → open the editor at `/d/<id>`. Browser back → My Files.
- Click "New blank" → editor opens at `/d/<new-id>`, server row exists,
  title editable in place.
- Ctrl+S → in-place save. **One** entry in My Files, not N.
- Click the logo → back to My Files, always.
- URL is canonical: refresh works, share the URL with yourself
  works, browser history works.

That's four files per repo for the critical path. ~8 total.

Everything else in §2–§4 stacks on top once the foundation exists.

---

## §7 — Open questions before any code

1. **Single mode definition** — confirm: single mode = one admin
   account, no peer-to-peer share-for-edit (only download / view-only).
   Multi mode = full account system + share-for-edit lands in phase 5.
2. **Router choice** — hash (`#/d/<id>`) vs pushState (`/d/<id>`).
   Hash needs no server change; pushState needs the server to serve
   `index.html` on every path (which it already does for `/admin`).
   **Recommend pushState.**
3. **"My Files" vs "My Spreadsheets" / "My Documents"** — language call.
   **Recommend "My Files"** because in Phase 5 we may share types
   (drag a docx into Sheets' file list? probably not, but the name
   leaves room).
4. **Templates location** — section on `/` vs `/templates` route.
   **Recommend section on `/`** for v1; promote to its own route only
   if the gallery grows beyond 12 entries.

Once §7 is answered I can start Phase 1 (sheet first, document
mirrors it). Won't push anything until then.
