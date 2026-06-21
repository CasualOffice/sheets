# Sharing Model — Casual Sheets

Design proposal for multi-mode sharing permissions. Branches from
`docs/UX_AUDIT.md` §2.9 (sheet) / §3.5 (doc). **No code yet** — this
doc exists so a future implementation has a definite shape to verify
against, instead of inventing semantics during a PR.

## §1 — What we already have today

Two sharing surfaces ship and work:

1. **`/r/<roomId>` anonymous coedit rooms** — built before personal-
   mode. The room URL is the capability: anyone with the link joins
   the same Y.Doc, picks an arbitrary display name, and edits. Optional
   password gate (`role=write`) lets the room creator restrict the
   write-role to recipients who know the secret. Read-only joiners
   get `role=view`.
2. **Personal-mode accounts (Phase C)** — single or multi mode.
   Single = one admin account, no peer accounts. Multi = an
   administrator can create per-user accounts. Files are scoped per
   user (`local.PerUserStores`).

The two systems were designed in isolation. The audit calls out the
gap: a multi-mode operator with five staff accounts has no way to say
"share this workbook with Alice for editing and Bob for view-only."
Today they have two bad options: send a `/r/<roomId>` link (anonymous,
no audit trail, anyone-with-link semantics) or paste the file via the
out-of-band channel.

## §2 — Hard constraints (decisions already made)

These are pinned by user direction and architecture; the design
respects them.

- **Single mode allows anonymous share-with-password.** Quote from
  the audit thread: _"for anonymous writer it's user's fault as he
  can share write access with password — architectural choice."_
  So single mode does NOT block anonymous edit; the room URL +
  password pattern stays as-is. Memory: `feedback-single-mode-share`.
- **Backend remains stateless.** The room manager keeps no user
  table; the Y.Doc is the source of truth for the live session.
  Any permission model has to encode authority in the join handshake,
  not in long-lived server-side ACLs in the gateway. Persistence
  for cross-session ACLs lives in the file host (`host.Integration`
  impl), not the gateway.
- **`docID = base64url(wopiSrc)`** for the WOPI mode. WOPI hosts own
  permissions; nothing in our model overrides them. The model below
  applies to **personal-mode** files only (Mode 3) — Mode 2 (WOPI) and
  Mode 1 (Pages demo) are untouched.

## §3 — Model

### §3.1 Identities

| Identity      | Auth source                       | Multi mode      | Single mode     |
| ------------- | --------------------------------- | --------------- | --------------- |
| **Admin**     | personal-mode login (admin role)  | full ACL writer | full ACL writer |
| **Member**    | personal-mode login (user role)   | per-share ACL   | n/a             |
| **Anonymous** | `?share=<token>` link, no account | per-link ACL    | per-link ACL    |

### §3.2 Roles per share

Three roles, single source of truth:

- **`view`** — read-only. Sees the workbook, can copy values out,
  cannot edit, cannot resolve comments, cannot follow filtered
  links to other peers' cursors.
- **`comment`** — read + add/resolve comments. No formula/value
  edits.
- **`edit`** — full Y.Doc participation. Edits, comment threads,
  rename, format. Cannot delete the file or change shares (admin-
  only).

(Suggestion mode like Google Docs is **out of scope for v1.**
Tracked as a separate proposal.)

### §3.3 Share artefacts

Two distinct artefacts per workbook, both persisted by the file host:

1. **Member ACLs** — `(workbookId, memberId, role)` rows. Multi-mode
   only. Bound to the member's account.
2. **Link tokens** — `(workbookId, token, role, expiresAt?, passwordHash?)`
   rows. Both modes. The token IS the capability; password is an
   optional layered gate.

Both shapes live in the file host. The gateway reads them only at
join-time to compute the joiner's effective role.

### §3.4 Join handshake

When a client opens `/r/<roomId>?share=<token>` or `/sheet/<id>`:

1. **Identity probe.** The gateway calls `host.WhoAmI(req)`. The
   host returns: `{ kind: 'member', id }`, `{ kind: 'anonymous' }`,
   or `{ kind: 'admin', id }`. (Today's `/auth/status` covers
   member + admin; anonymous is the implicit fallback.)
2. **Effective role.** The gateway computes the **highest** of:
   - For `admin` → always `edit`.
   - For `member` → `host.GetMemberRole(workbookId, memberId)` if
     a member ACL row exists; else fall through.
   - For any kind with a `?share=<token>` → `host.GetLinkRole(token)`
     after password verification.
   - Default = no access → 403.
3. **Y.Doc role enforcement.** Today's `applyViewOnlyMode` flips
   Univer permissions per-unit-id on view joiners. Extend it with a
   `comment` mode that locks Y maps + arrays but lets comment
   mutations through.

### §3.5 UX — Share dialog

Multi-mode share dialog ships **two tabs**:

```
┌─ Share "Q3 budget.xlsx" ──────────────────────────┐
│                                                   │
│  [Members]  [Link]                                │
│                                                   │
│  ┌─ Members tab ───────────────────────────────┐  │
│  │  Add member…           [picker]    Role: ▾ │  │
│  │  alice@…                          Edit  ✕  │  │
│  │  bob@…                            View  ✕  │  │
│  └────────────────────────────────────────────┘  │
│                                                   │
│  ┌─ Link tab ──────────────────────────────────┐  │
│  │  Anyone with the link can: ▾ View / Edit   │  │
│  │  Expires: ▾ Never / 7 days / 30 days       │  │
│  │  Password: ☐ Require a password to join    │  │
│  │  https://sheet.example.com/r/abc?share=xyz │  │
│  │  [Copy] [Reset link]                       │  │
│  └────────────────────────────────────────────┘  │
│                                                   │
│                                       [ Done ]    │
└───────────────────────────────────────────────────┘
```

Single mode shows the **Link tab only** — Members doesn't apply
when there are no peer accounts. WOPI mode hides the whole dialog;
the WOPI host owns sharing.

### §3.6 Audit trail

Every effective-role decision is logged through the existing
structured logger (`backend/internal/auth/personal` request-id
middleware, commit `a63941c`):

```
{
  "evt": "share.join",
  "workbookId": "...",
  "viaToken": "abc",   // or null
  "viaMember": "alice@...",  // or null
  "effectiveRole": "edit",
  "requestId": "..."
}
```

No UI for the log in v1 — the admin reads server logs. v2 surfaces
the join history per workbook in the share dialog ("Last 30 days:
alice@… (edit, 12×), bob@… (view, 3×)").

## §4 — Backend surface

New routes (multi mode only; single mode is link-only):

- `GET    /api/files/{id}/shares` — list ACL rows + link tokens.
- `POST   /api/files/{id}/shares/member` — `{ memberId, role }`.
- `PATCH  /api/files/{id}/shares/member/{memberId}` — `{ role }`.
- `DELETE /api/files/{id}/shares/member/{memberId}` — revoke.
- `POST   /api/files/{id}/shares/link` — mint a token. Body:
  `{ role, expiresInDays?, password? }`. Returns the URL fragment.
- `PATCH  /api/files/{id}/shares/link/{token}` — flip role / expiry.
- `DELETE /api/files/{id}/shares/link/{token}` — revoke.

All routes gate on `RequireAdmin` or `RequireOwner(workbookId)`
(host returns the owner from the file metadata, which already
tracks `userID` per file under personal mode).

The `host.Locker` capability (Phase D #4, commit `bfc5e4b`) keeps
working: the room manager claims the host lock on first join,
releases on drain. The join handshake's role check is layered ON
TOP of locking — a member with edit role still races for the lock
the normal way.

## §5 — Open trade-offs (decide before implementing)

1. **Anonymous edit in single mode.** Pinned per user direction.
   But the multi-mode link tab also exposes anyone-with-link-edit
   tokens. We accept this — multi-mode operators are also explicit
   about the choice.
2. **Email vs username for members.** Existing accounts use
   `username`; the share dialog could expose either. Industry
   default is email. **Recommend** display email when it exists,
   fall back to username — the existing `Profile` shape (commit
   `2e197d7`) has both.
3. **Suggestion mode.** Out of scope; tracked separately. Adding
   it later means a fourth role with a different Y.Doc enforcement
   path (revisions feed of pending edits).
4. **Cross-workbook share defaults.** Industry tools cache "last
   shared with…" per-account. v1 has no caching; v2 can revisit.
5. **Revocation latency.** A revoke today only kicks in on the
   NEXT join. Already-connected clients keep their session until
   their WS drops. v2 may add a server-side broadcast that boots
   stale roles in-session.

## §6 — Phasing

| Phase                               | Scope                                                                                         | Effort    | Status                                       |
| ----------------------------------- | --------------------------------------------------------------------------------------------- | --------- | -------------------------------------------- |
| §6.1 — Link tokens (single + multi) | `POST /share/link` + token role enforcement in the join handshake + Link tab in share dialog. | 1–2 weeks | ✅ **done**                                  |
| §6.2 — Member ACLs (multi only)     | Routes + Members tab + email lookup.                                                          | 1 week    | 🟡 foundation done; enforcement + UI pending |
| §6.3 — Audit logging surface        | Structured logs + admin-side log view.                                                        | 3 days    | ⬜                                           |
| §6.4 — Suggestion mode (proposal)   | Out-of-scope — separate doc.                                                                  | TBD       | ⬜                                           |

§6.1 alone closes the immediate "no recipient permission story"
gap. §6.2 is the multi-mode follow-up. §6.3 makes the operator-side
story complete.

### Implementation status (2026-06-21)

**§6.1 — done.** Link tokens are room-bound `(workbookId, roomId, token,
role, expiresAt?, passwordHash?)` rows in the personal SQLite store
(`apps/server/src/auth/personal.ts`); owner/admin-gated CRUD at
`/files/:id/shares[/link/:token]`; a public `GET /files/shares/link/:token/meta`
for pre-join discovery (`{valid, role, hasPassword, roomId}` — never the hash).
Enforcement is the **server-authoritative, pure** `resolveJoinRole`
(`apps/server/src/auth/join-role.ts`) wired into the Hocuspocus
`onAuthenticate` (`yjs.ts`): a `?share=<token>` makes the server decide the
role (client `?role=` ignored); the token must resolve, be bound to the room
being joined (replay-proof), and pass its optional `?sp=` bcrypt password. No
token → the legacy anonymous path is byte-identical. Client: a secure-link
section in the share dialog (mint view/edit, optional expiry + password, list,
copy, revoke) plus join-side `?share=`/`?sp=` forwarding + a password prompt.

- **Deferred from §6.1:** the `comment` role currently collapses to read-only
  (Hocuspocus `readOnly` is binary; fine-grained comment-mode — permit comment
  mutations, block edits — needs Univer-permission work in the client's
  `applyViewOnlyMode`, per §3.4).

**§6.2 — foundation done, enforcement pending.** Member-ACL persistence
(`file_member_acls`, upsert + `getMemberRole` + `findMemberByHandle`) and the
multi-mode owner/admin-gated routes (`/files/:id/shares/member[s]`) are in.
**Not yet built:** (a) join enforcement for tokenless member joins — needs a
**room→workbook mapping** so a logged-in member opening `/r/<roomId>` resolves
to the workbook whose ACL governs them (open design decision: a dedicated
`room_workbooks` association recorded when a room is created for a file, vs.
deriving it from an existing `share_links` row for that room); (b) the
"Shared with me" file surface + member room-join flow; (c) the Members tab UI.
Until enforcement lands, member ACLs are stored but grant no live access.

## §7 — What this proposal does **not** do

- Does not change the `/r/<roomId>` URL shape. Existing share-link
  links keep working.
- Does not invent a new file host capability. Everything sits on
  the existing `host.Integration` + `host.Locker` surfaces, with
  one new sub-interface `host.Shares` for ACL persistence.
- Does not touch WOPI. WOPI hosts own permissions; we honour their
  decision.
- Does not add per-cell or per-range ACLs. Workbook-level only.

## §8 — Open questions for the next review

1. Should single mode allow Member ACLs (between the admin and
   anonymous-with-token)? Default **no** — single = one account.
2. Should expired link tokens auto-purge or keep history? Default
   **keep**, hidden behind a "show expired" toggle.
3. Should we expose API tokens for programmatic shares (CI / scripts
   uploading a workbook with a share URL)? Out of scope for v1.
