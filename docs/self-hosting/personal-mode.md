# Personal mode (Mode 3)

> Phase C of the storage-modes work (#49) — landed 2026-06-06.

Standalone docker with personal accounts. Sign up once, your files persist
to a volume, the same files reopen on every visit and from any machine
pointed at the same server.

This is Joel's case from #48: "save files to a container volume so the
last file used is open." It's not multi-tenant SaaS — the next step
(workspaces + RBAC + invites) lives in a separate tracker (team mode).

For the architectural design + the broader 3-mode picture see
[`docs/STORAGE_MODES.md`](../STORAGE_MODES.md).

---

## TL;DR

```yaml
# docker-compose.yml
services:
  app:
    image: schnsrw/casual-sheets:0.2
    ports: ['3000:3000']
    environment:
      CASUAL_PERSONAL_MODE: single # 'single' (one account) or 'multi'
      CASUAL_STORAGE: local
      CASUAL_LOCAL_PATH: /data/workbooks
    volumes:
      - ./data:/data
```

`docker compose up -d`, open `http://localhost:3000`, sign up. Your
account is the admin. Files land in `./data/workbooks/`; the users + sessions
table lives in `./data/users.db`.

---

## Configuration

| Env var                 | Default           | What it does                                                                                                                                                                                                                                                                                     |
| ----------------------- | ----------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `CASUAL_PERSONAL_MODE`  | `none`            | `none` (off — WOPI-only), `single` (one account; signup closes after the first user), `multi` (open signup; first user is admin)                                                                                                                                                                 |
| `CASUAL_USERS_DB_PATH`  | `/data/users.db`  | SQLite file holding users, sessions, files registry, and profile blobs. Survives restarts via the volume.                                                                                                                                                                                        |
| `CASUAL_BOOTSTRAP_USER` | _unset_           | One-shot — when the users table is empty, creates an admin from `<username>:<password>`. Useful for upgrades from a pre-Phase-C install so the operator doesn't have to click through signup. Quietly ignored once any user exists; rewriting the env var does **not** rotate the live password. |
| `CASUAL_LOCAL_PATH`     | `/data/workbooks` | Where the `local` host backend stores the actual `.xlsx` bytes. (Same env as the existing WOPI flow.)                                                                                                                                                                                            |
| `MAX_UPLOAD_MB`         | `100`             | Cap for uploads (server-side). The web client also has a `VITE_MAX_OPEN_MB` baked at build time — keep them in sync if you raise either.                                                                                                                                                         |
| `CASUAL_STORAGE`        | `memory`          | `local` / `s3` / `postgres` for the file bytes. Same setting that's already documented in [`docs/ENV.md`](../ENV.md); personal mode layers on top, doesn't replace.                                                                                                                              |

`CASUAL_PERSONAL_MODE=none` is the **default** so existing WOPI / anonymous
deployments are not affected by an upgrade. Operators flip personal mode
on explicitly.

---

## What lives on the volume

```
/data/
├── users.db                # SQLite — users + sessions + files registry + profile
├── users.db-wal            # SQLite WAL — leave it
├── users.db-shm            # SQLite shared-memory — leave it
├── casual-admin.json       # Admin-panel config (pre-existing)
└── workbooks/
    ├── f-3b8a9d4e2c7a.xlsx       # Workbook bytes, keyed by registry id
    ├── f-3b8a9d4e2c7a.meta.json  # Per-file metadata sidecar
    └── …
```

Backups: tar the entire `/data/` tree. The SQLite file is WAL-mode; safe to
copy live, but for paranoid backups run `docker exec <container>
sqlite3 /data/users.db ".backup '/data/users-snap.db'"` first.

---

## Sign-up + sign-in flow

First visit on a fresh `single` install:

1. Server is in `single` mode + zero users → **signup screen** (the
   "first account is the admin" copy)
2. User fills username + password (min 8 chars) → POST `/auth/signup` →
   201 + `cs_session` cookie set
3. Subsequent visits → cookie still valid → straight into the editor

`multi` keeps signup open after the first user; every additional user
is a plain (non-admin) account. Per-user file isolation is total — an
admin **cannot** list or open another user's workbook through the UI
or the API. (Team mode changes this.)

---

## Profile + preferences

Click the avatar in the title bar → **Settings**. Three tabs:

- **Profile** — display name, optional email, timezone, avatar
  (PNG/JPEG/WebP/GIF, ≤ 256 KB; resized client-side before upload).
- **Security** — change password (every other session for the user is
  blown out so a stolen browser stops working immediately). Danger
  zone: delete account (refuses for the last admin).
- **Preferences** — theme, language, date format. Stored as a JSON
  blob the server doesn't interpret — client owns the meaning.

Profile fields are optional everywhere. A freshly signed-up user has
no email, no timezone (defaults to `UTC`), no avatar; the UI shows the
username as the display name.

---

## Resetting a forgotten password

No SMTP in v1 — the design recommendation lands fine here (decided
2026-06-06): homelab and airgapped users don't typically have an MTA,
and `docker exec` is documented enough to be the recovery path.

```sh
# Interactive — prompts on stdin with echo disabled
docker exec -it <container> casual-sheets-reset-password <username>

# Non-interactive (CI / scripted)
docker exec <container> casual-sheets-reset-password <username> --password='<new>'
```

Behaviour:

- Reads `CASUAL_USERS_DB_PATH` (default `/data/users.db`).
- All active sessions for the user are invalidated — every other tab
  and machine gets logged out on next request.
- Refuses passwords shorter than 8 characters.
- Exit codes: `0` ok, `1` user not found, `2` password rejected, `3`
  couldn't open the DB, `4` bad invocation.

Running locally without docker:

```sh
CASUAL_USERS_DB_PATH=/path/to/users.db \
  pnpm --filter @sheet/server reset-password <username>
```

---

## Migration from pre-Phase-C deployments

A WOPI-only install pre-dating Phase C has files under `/data/workbooks/`
but no `/data/users.db` and no users.

Promotion path:

1. **Decide on a bootstrap admin.** Pick a username + password.
2. **Set the env**:

   ```sh
   CASUAL_PERSONAL_MODE=single
   CASUAL_BOOTSTRAP_USER=admin:strongpassword
   ```

3. **Restart** the container. On first boot the bootstrap admin is
   created in the new `users.db`. The orphan `.xlsx` files in
   `/data/workbooks/` remain on disk but are **not** automatically
   adopted into the new admin's My Files list — adoption requires a
   manual one-time pass (forthcoming follow-up; tracked as a Mode 3
   stretch goal in #49).

4. **Unset `CASUAL_BOOTSTRAP_USER` after first boot.** The store
   silently no-ops the bootstrap when any user exists — but leaving
   it in the env is documented as "no effect" rather than "secret".
   Cleaner to remove it.

---

## What's _not_ in Mode 3

- **Sharing files between users.** Each user only sees their own
  workbooks. The existing co-edit room URL flow still works for
  anonymous ephemeral collab between any two browsers.
- **Email-based password recovery.** CLI reset is the documented
  path; SMTP plumbing arrives with team mode.
- **Workspaces, role-based access management, invites, audit log.**
  All on the team-mode roadmap (separate tracker; closes a different
  set of acceptance criteria).
- **Folder tree.** Flat listing in v1. Folders if and when users ask.
- **Server-side version history.** Per-workbook history still lives
  in browser IDB. Server-side versioning waits for a real ask.

---

## Acceptance — what "personal mode shipped" looks like

A fresh `docker run` with `CASUAL_PERSONAL_MODE=single` + a `/data`
volume + a browser produces:

- [x] Welcome screen → create admin account → land in My Files (empty)
- [x] Upload `.xlsx` from disk → file appears in My Files
- [x] Open → edit → File → Save → toast: "Saved to `<filename>`"
- [x] Browser reload → cookie still valid → My Files still lists the
      file → opening it shows the edited content
- [x] Sign out from the AccountMenu → login screen → log in → file
      still there
- [x] Settings: change display name, set timezone, upload avatar; new
      avatar shows in the title-bar circle and across browser tabs
- [x] Change password → all other tabs forced to re-login
- [x] CLI reset from `docker exec` → old password rejected, new one
      accepted

All verified locally on 2026-06-06; CI-level coverage tracked in #49.
