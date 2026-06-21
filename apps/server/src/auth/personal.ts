import Database from 'better-sqlite3';
import bcrypt from 'bcryptjs';
import { randomBytes } from 'node:crypto';
import { mkdirSync } from 'node:fs';
import { dirname } from 'node:path';

/**
 * Personal-mode auth store (Phase C of the storage-modes work, #49).
 *
 * One SQLite database — by default `/data/users.db` — holds the
 * users + sessions tables. SQLite is the right shape here: the
 * standalone docker image is single-process by design (Mode 3 is
 * personal, not multi-tenant SaaS), and avoiding a Postgres dep
 * keeps the deploy story to "one container + one volume."
 *
 * Modes
 *
 *   - `none`  — the auth tables exist but `/auth/signup` and
 *               `/auth/login` are disabled (503). The server behaves
 *               exactly like today's WOPI-only / anonymous deploy.
 *   - `single`— first signup creates the admin; further signups are
 *               rejected (403). Designed for "I want my own files in
 *               my own docker."
 *   - `multi` — open signup. First account is the admin (config +
 *               room limits). Per-user file isolation is total — even
 *               an admin cannot list or open another user's workbook.
 *
 * Bootstrap
 *
 *   `CASUAL_BOOTSTRAP_USER=<username>:<password>` on first server
 *   start (when the users table is empty) creates that user as the
 *   admin. Useful for upgraders coming from a pre-Phase-C install
 *   who already have files in `/data/workbooks` and don't want to
 *   click through the signup screen. Wiped to a no-op after the
 *   account exists — re-setting the env var doesn't change a live
 *   password.
 *
 * Password recovery
 *
 *   No SMTP in v1 — the operator runs `casual-sheets reset-password
 *   <username>` from a docker exec shell. See Phase C Batch 5 for the
 *   CLI implementation.
 */

export type PersonalMode = 'none' | 'single' | 'multi';

export type PublicUser = {
  id: number;
  username: string;
  isAdmin: boolean;
  createdAt: number;
};

/** Editable profile fields. Optional everywhere — a freshly-signed-up
 *  user has none of these set, and the UI defaults the display name
 *  to the username. */
export type UserProfile = {
  displayName: string | null;
  email: string | null;
  /** IANA time-zone (e.g. "Europe/Berlin"). Defaults to "UTC". */
  timezone: string;
  /** Truthy when an avatar blob has been uploaded — the route layer
   *  serves the bytes from `GET /auth/profile/avatar`. */
  hasAvatar: boolean;
  /** Free-form JSON for client preferences (theme, language, date
   *  format, etc.). Server doesn't interpret — it just round-trips. */
  preferences: Record<string, unknown>;
};

export type FullUser = PublicUser & UserProfile;

/** Row in the personal-mode files registry. The byte payload lives in
 *  the `HostIntegration` backend keyed by `id`; this record is the
 *  ownership + metadata pointer. */
export type FileRecord = {
  id: string;
  ownerId: number;
  displayName: string;
  size: number;
  /** Opaque version, bumped on every write. Matches the WOPI
   *  `Version` field, used for `If-Match` conflict detection. */
  etag: string;
  createdAt: number;
  modifiedAt: number;
};

/** Share-link role. The token IS the capability; the role decides how
 *  much it grants. Mirrors `docs/SHARING_MODEL.md` §3.2. */
export type ShareRole = 'view' | 'comment' | 'edit';

/** A persisted link token (sharing-model §3.3). The token string is
 *  the capability secret; `passwordHash` (when set) is an optional
 *  layered gate the joiner must also satisfy. */
export type ShareLink = {
  workbookId: string;
  token: string;
  role: ShareRole;
  /** The collab room (Yjs `documentName`) this token is bound to at
   *  mint time. Rooms are anonymous and NOT keyed by workbookId, so a
   *  token scoped only to a workbook could be replayed on ANY room.
   *  Binding to a specific roomId closes that replay gap — the
   *  enforcement gate rejects a token whose roomId doesn't match the
   *  room being joined. See `resolveJoinRole`. */
  roomId: string;
  /** ms epoch, or null for a never-expiring link. */
  expiresAt: number | null;
  /** bcrypt hash of the optional join password, or null. Never leaves
   *  the store — the route layer surfaces only `hasPassword`. */
  passwordHash: string | null;
  createdAt: number;
  /** `users.id` of the admin/owner who minted the token. */
  createdBy: number;
};

/** Read shape for the future join-handshake — the minimum a caller
 *  needs to compute an effective role. `getLinkRole` returns this and
 *  bakes in expiry (null when the token has lapsed). */
export type ShareLinkRole = {
  workbookId: string;
  /** The room this token is bound to. The enforcement gate compares
   *  this against the `documentName` being joined and rejects on a
   *  mismatch — a token minted for room A grants nothing on room B. */
  roomId: string;
  role: ShareRole;
  hasPassword: boolean;
  passwordHash: string | null;
  expiresAt: number | null;
};

export type PersonalAuthOptions = {
  dbPath: string;
  mode: PersonalMode;
  /** Bootstrap admin spec — "username:password". Applied once when the
   *  users table is empty. */
  bootstrap: string | null;
  /** Session lifetime in ms. Defaults to 30 days. */
  sessionTtlMs?: number;
};

export type CreateUserResult =
  | { ok: true; user: PublicUser }
  | {
      ok: false;
      reason:
        | 'username-taken'
        | 'mode-disabled'
        | 'signup-closed'
        | 'weak-password'
        | 'invalid-username';
    };

export type LoginResult =
  | { ok: true; user: PublicUser }
  | { ok: false; reason: 'mode-disabled' | 'invalid-credentials' };

const BCRYPT_ROUNDS = 10;
const DEFAULT_SESSION_TTL_MS = 30 * 24 * 60 * 60 * 1000;
const USERNAME_RE = /^[a-zA-Z0-9_.-]{2,40}$/;
const MIN_PASSWORD_LEN = 8;
/** ~32 bytes of CSPRNG entropy, base64url-encoded — the link token. */
const SHARE_TOKEN_BYTES = 32;
const SHARE_ROLES: readonly ShareRole[] = ['view', 'comment', 'edit'];

/** Type-guard for the role enum — used at the route boundary. */
export function isShareRole(v: unknown): v is ShareRole {
  return typeof v === 'string' && (SHARE_ROLES as readonly string[]).includes(v);
}

/**
 * Thin wrapper around a single SQLite connection. The Fastify app
 * holds one instance for the process lifetime; serialised IO is fine
 * at Mode 3's scale (single-user docker, occasionally a small team
 * in `multi`).
 */
export class PersonalAuthStore {
  readonly mode: PersonalMode;
  private readonly db: Database.Database;
  private readonly sessionTtlMs: number;

  constructor(opts: PersonalAuthOptions) {
    this.mode = opts.mode;
    this.sessionTtlMs = opts.sessionTtlMs ?? DEFAULT_SESSION_TTL_MS;

    // Ensure the parent dir exists — by default `/data` should already
    // be a volume, but a dev launch might point this elsewhere.
    mkdirSync(dirname(opts.dbPath), { recursive: true });

    this.db = new Database(opts.dbPath);
    this.db.pragma('journal_mode = WAL');
    this.db.pragma('foreign_keys = ON');
    this.migrate();

    if (opts.bootstrap && this.countUsers() === 0) {
      this.applyBootstrap(opts.bootstrap);
    }
  }

  // ── Public API ─────────────────────────────────────────────────────────

  countUsers(): number {
    return (this.db.prepare('SELECT COUNT(*) AS n FROM users').get() as { n: number }).n;
  }

  hasAnyUser(): boolean {
    return this.countUsers() > 0;
  }

  /** Used by routes to decide whether `POST /auth/signup` is open. */
  signupAllowed(): boolean {
    if (this.mode === 'none') return false;
    if (this.mode === 'multi') return true;
    // single: only allowed before the first user exists.
    return !this.hasAnyUser();
  }

  createUser(username: string, password: string): CreateUserResult {
    if (this.mode === 'none') return { ok: false, reason: 'mode-disabled' };
    if (this.mode === 'single' && this.hasAnyUser()) {
      return { ok: false, reason: 'signup-closed' };
    }
    if (!USERNAME_RE.test(username)) {
      return { ok: false, reason: 'invalid-username' };
    }
    if (password.length < MIN_PASSWORD_LEN) {
      return { ok: false, reason: 'weak-password' };
    }
    const existing = this.db
      .prepare('SELECT id FROM users WHERE username = ? COLLATE NOCASE')
      .get(username);
    if (existing) return { ok: false, reason: 'username-taken' };

    const hash = bcrypt.hashSync(password, BCRYPT_ROUNDS);
    const now = Date.now();
    // First user is admin. Beyond that, plain users.
    const isAdmin = this.countUsers() === 0 ? 1 : 0;
    const result = this.db
      .prepare(
        'INSERT INTO users (username, password_hash, is_admin, created_at) VALUES (?, ?, ?, ?)',
      )
      .run(username, hash, isAdmin, now);
    const id = Number(result.lastInsertRowid);
    return {
      ok: true,
      user: { id, username, isAdmin: isAdmin === 1, createdAt: now },
    };
  }

  verifyLogin(username: string, password: string): LoginResult {
    if (this.mode === 'none') return { ok: false, reason: 'mode-disabled' };
    const row = this.db
      .prepare(
        'SELECT id, username, password_hash, is_admin, created_at FROM users WHERE username = ? COLLATE NOCASE',
      )
      .get(username) as
      | {
          id: number;
          username: string;
          password_hash: string;
          is_admin: number;
          created_at: number;
        }
      | undefined;
    if (!row) return { ok: false, reason: 'invalid-credentials' };
    if (!bcrypt.compareSync(password, row.password_hash)) {
      return { ok: false, reason: 'invalid-credentials' };
    }
    return {
      ok: true,
      user: {
        id: row.id,
        username: row.username,
        isAdmin: row.is_admin === 1,
        createdAt: row.created_at,
      },
    };
  }

  /** Mint a fresh session for the user. Returns the opaque session id
   *  the client stores in the `cs_session` cookie. */
  startSession(userId: number): { sessionId: string; expiresAt: number } {
    const sessionId = randomBytes(24).toString('hex');
    const expiresAt = Date.now() + this.sessionTtlMs;
    this.db
      .prepare('INSERT INTO sessions (id, user_id, expires_at) VALUES (?, ?, ?)')
      .run(sessionId, userId, expiresAt);
    return { sessionId, expiresAt };
  }

  /** Resolve a session id to a user. Slides the expiry forward on a
   *  hit (rolling sessions). Returns null for unknown / expired / GCd. */
  resolveSession(sessionId: string | null | undefined): PublicUser | null {
    if (!sessionId) return null;
    const row = this.db
      .prepare(
        `SELECT u.id, u.username, u.is_admin, u.created_at, s.expires_at
           FROM sessions s JOIN users u ON u.id = s.user_id
          WHERE s.id = ?`,
      )
      .get(sessionId) as
      | { id: number; username: string; is_admin: number; created_at: number; expires_at: number }
      | undefined;
    if (!row) return null;
    const now = Date.now();
    if (row.expires_at < now) {
      this.db.prepare('DELETE FROM sessions WHERE id = ?').run(sessionId);
      return null;
    }
    // Slide expiry forward — keeps the user signed in as long as
    // they're active without re-issuing on every request.
    const newExpires = now + this.sessionTtlMs;
    if (newExpires - row.expires_at > 60_000) {
      this.db.prepare('UPDATE sessions SET expires_at = ? WHERE id = ?').run(newExpires, sessionId);
    }
    return {
      id: row.id,
      username: row.username,
      isAdmin: row.is_admin === 1,
      createdAt: row.created_at,
    };
  }

  endSession(sessionId: string): void {
    this.db.prepare('DELETE FROM sessions WHERE id = ?').run(sessionId);
  }

  /** Pruning pass — kills expired session rows. Cheap full-table scan
   *  via the `expires_at` index. */
  pruneExpiredSessions(): number {
    const result = this.db.prepare('DELETE FROM sessions WHERE expires_at < ?').run(Date.now());
    return result.changes;
  }

  /** Change a user's password — used by the account modal + by the
   *  CLI reset subcommand. Returns false if the current password
   *  check fails (skipped when `currentPassword === null`, the
   *  CLI's escape hatch). */
  changePassword(userId: number, currentPassword: string | null, newPassword: string): boolean {
    if (newPassword.length < MIN_PASSWORD_LEN) return false;
    const row = this.db.prepare('SELECT password_hash FROM users WHERE id = ?').get(userId) as
      | { password_hash: string }
      | undefined;
    if (!row) return false;
    if (currentPassword !== null && !bcrypt.compareSync(currentPassword, row.password_hash)) {
      return false;
    }
    const hash = bcrypt.hashSync(newPassword, BCRYPT_ROUNDS);
    this.db.prepare('UPDATE users SET password_hash = ? WHERE id = ?').run(hash, userId);
    // Force re-login on every other session — protects against a
    // compromised browser surviving a deliberate password change.
    this.db.prepare('DELETE FROM sessions WHERE user_id = ?').run(userId);
    return true;
  }

  /** Delete a user + their sessions. Refuses to delete the last
   *  admin — the operator would lock themselves out of the admin
   *  panel. */
  deleteUser(userId: number): boolean {
    const target = this.db.prepare('SELECT is_admin FROM users WHERE id = ?').get(userId) as
      | { is_admin: number }
      | undefined;
    if (!target) return false;
    if (target.is_admin === 1) {
      const otherAdmins = this.db
        .prepare('SELECT COUNT(*) AS n FROM users WHERE is_admin = 1 AND id != ?')
        .get(userId) as { n: number };
      if (otherAdmins.n === 0) return false;
    }
    this.db.prepare('DELETE FROM users WHERE id = ?').run(userId);
    return true;
  }

  // ── Profile (display name / email / timezone / avatar / preferences) ──

  /** Read the editable profile for a signed-in user. */
  getProfile(userId: number): UserProfile | null {
    const row = this.db
      .prepare(
        `SELECT display_name, email, timezone, avatar_mime, preferences
           FROM users WHERE id = ?`,
      )
      .get(userId) as
      | {
          display_name: string | null;
          email: string | null;
          timezone: string;
          avatar_mime: string | null;
          preferences: string;
        }
      | undefined;
    if (!row) return null;
    let prefs: Record<string, unknown> = {};
    try {
      prefs = JSON.parse(row.preferences ?? '{}');
    } catch {
      // Corrupt JSON — treat as empty so the route doesn't 500.
    }
    return {
      displayName: row.display_name,
      email: row.email,
      timezone: row.timezone ?? 'UTC',
      hasAvatar: Boolean(row.avatar_mime),
      preferences: prefs,
    };
  }

  /** Update one or more profile fields. Each field is optional — only
   *  the keys present in `patch` are written. Returns null when the
   *  user is unknown OR on a uniqueness collision (`email`). */
  updateProfile(
    userId: number,
    patch: Partial<Pick<UserProfile, 'displayName' | 'email' | 'timezone' | 'preferences'>>,
  ): UserProfile | null {
    // Capacity: validate before mutating. Email goes through the
    // unique index on email; trim + lowercase normalisation isn't
    // server-imposed (COLLATE NOCASE is on the column).
    if (patch.displayName !== undefined) {
      const dn = patch.displayName?.trim() ?? '';
      if (dn.length > 80) return null;
    }
    if (patch.email !== undefined && patch.email !== null) {
      const em = patch.email.trim();
      // Cheap shape check — proper validation is a UI concern (the
      // input element does most of the lifting).
      if (em && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em)) return null;
    }
    if (patch.timezone !== undefined) {
      // IANA tzs only — validated by `Intl.supportedValuesOf('timeZone')`
      // on the client; the server treats it as an opaque string but
      // sanity-checks the length.
      if (patch.timezone.length > 64) return null;
    }

    const fields: string[] = [];
    const values: Array<unknown> = [];
    if (patch.displayName !== undefined) {
      fields.push('display_name = ?');
      values.push(patch.displayName?.trim() || null);
    }
    if (patch.email !== undefined) {
      fields.push('email = ?');
      const em = patch.email?.trim();
      values.push(em ? em : null);
    }
    if (patch.timezone !== undefined) {
      fields.push('timezone = ?');
      values.push(patch.timezone.trim() || 'UTC');
    }
    if (patch.preferences !== undefined) {
      fields.push('preferences = ?');
      values.push(JSON.stringify(patch.preferences));
    }
    if (fields.length === 0) return this.getProfile(userId);

    try {
      this.db.prepare(`UPDATE users SET ${fields.join(', ')} WHERE id = ?`).run(...values, userId);
    } catch (err) {
      // Unique-email collision lands here.
      if ((err as { code?: string }).code === 'SQLITE_CONSTRAINT_UNIQUE') return null;
      throw err;
    }
    return this.getProfile(userId);
  }

  /** Write an avatar blob. Caller is expected to cap the size + mime
   *  upstream; the store doesn't second-guess. Pass `null` to clear. */
  setAvatar(userId: number, mime: string | null, bytes: Uint8Array | null): void {
    if (mime === null || bytes === null) {
      this.db
        .prepare('UPDATE users SET avatar_mime = NULL, avatar_blob = NULL WHERE id = ?')
        .run(userId);
      return;
    }
    this.db
      .prepare('UPDATE users SET avatar_mime = ?, avatar_blob = ? WHERE id = ?')
      .run(mime, Buffer.from(bytes), userId);
  }

  /** Read the avatar bytes. Returns null when no avatar is set. */
  getAvatar(userId: number): { mime: string; bytes: Uint8Array } | null {
    const row = this.db
      .prepare('SELECT avatar_mime, avatar_blob FROM users WHERE id = ?')
      .get(userId) as { avatar_mime: string | null; avatar_blob: Buffer | null } | undefined;
    if (!row?.avatar_mime || !row?.avatar_blob) return null;
    return { mime: row.avatar_mime, bytes: new Uint8Array(row.avatar_blob) };
  }

  /** Resolve a username back to its id — used by the CLI reset. */
  findIdByUsername(username: string): number | null {
    const row = this.db
      .prepare('SELECT id FROM users WHERE username = ? COLLATE NOCASE')
      .get(username) as { id: number } | undefined;
    return row?.id ?? null;
  }

  /** Diagnostic — total users, sessions, db size on disk. */
  stats(): { userCount: number; sessionCount: number } {
    return {
      userCount: this.countUsers(),
      sessionCount: (this.db.prepare('SELECT COUNT(*) AS n FROM sessions').get() as { n: number })
        .n,
    };
  }

  close(): void {
    this.db.close();
  }

  // ── Files registry ─────────────────────────────────────────────────────

  /** Generate a fresh opaque file id. 16 hex chars + `f-` prefix —
   *  enough entropy for a personal-mode docker (the owner_id FK is
   *  the real ownership boundary; the id is just a server-issued
   *  handle the web client tosses around). */
  static newFileId(): string {
    return `f-${randomBytes(8).toString('hex')}`;
  }

  /** Insert a new file row. Returns the assigned id (the caller is
   *  expected to use it as the host integration's `FileId` for the
   *  byte payload). */
  createFile(opts: {
    ownerId: number;
    displayName: string;
    size: number;
    etag: string;
  }): FileRecord {
    const id = PersonalAuthStore.newFileId();
    const now = Date.now();
    this.db
      .prepare(
        `INSERT INTO files (id, owner_id, display_name, size, etag, created_at, modified_at)
         VALUES (?, ?, ?, ?, ?, ?, ?)`,
      )
      .run(id, opts.ownerId, opts.displayName, opts.size, opts.etag, now, now);
    return {
      id,
      ownerId: opts.ownerId,
      displayName: opts.displayName,
      size: opts.size,
      etag: opts.etag,
      createdAt: now,
      modifiedAt: now,
    };
  }

  /** Read a single file row. Returns null when the id is unknown. */
  getFile(id: string): FileRecord | null {
    const row = this.db
      .prepare(
        `SELECT id, owner_id, display_name, size, etag, created_at, modified_at
           FROM files WHERE id = ?`,
      )
      .get(id) as
      | {
          id: string;
          owner_id: number;
          display_name: string;
          size: number;
          etag: string;
          created_at: number;
          modified_at: number;
        }
      | undefined;
    if (!row) return null;
    return {
      id: row.id,
      ownerId: row.owner_id,
      displayName: row.display_name,
      size: row.size,
      etag: row.etag,
      createdAt: row.created_at,
      modifiedAt: row.modified_at,
    };
  }

  /** All files belonging to `ownerId`, newest-edit first. The route
   *  layer surfaces this to the web client's recents list. */
  listFilesForUser(ownerId: number): FileRecord[] {
    const rows = this.db
      .prepare(
        `SELECT id, owner_id, display_name, size, etag, created_at, modified_at
           FROM files WHERE owner_id = ? ORDER BY modified_at DESC`,
      )
      .all(ownerId) as Array<{
      id: string;
      owner_id: number;
      display_name: string;
      size: number;
      etag: string;
      created_at: number;
      modified_at: number;
    }>;
    return rows.map((row) => ({
      id: row.id,
      ownerId: row.owner_id,
      displayName: row.display_name,
      size: row.size,
      etag: row.etag,
      createdAt: row.created_at,
      modifiedAt: row.modified_at,
    }));
  }

  /** Bump size + etag + modified-at after a put. Caller has already
   *  written the bytes via the host integration. */
  recordFileUpdate(id: string, size: number, etag: string): void {
    this.db
      .prepare('UPDATE files SET size = ?, etag = ?, modified_at = ? WHERE id = ?')
      .run(size, etag, Date.now(), id);
  }

  /** Change a file's display name. */
  renameFile(id: string, newName: string): boolean {
    if (!newName.trim()) return false;
    const result = this.db
      .prepare('UPDATE files SET display_name = ?, modified_at = ? WHERE id = ?')
      .run(newName.trim(), Date.now(), id);
    return result.changes > 0;
  }

  /** Drop a file row. Caller is expected to also drop the bytes from
   *  the host integration. Returns false if the id was unknown. */
  deleteFile(id: string): boolean {
    const result = this.db.prepare('DELETE FROM files WHERE id = ?').run(id);
    return result.changes > 0;
  }

  // ── Share links (sharing-model §6.1 — persistence only) ─────────────────
  //
  // NOTE: these rows are inert until the join-handshake batch wires
  // `getLinkRole` into the room manager. Minting a token grants NO
  // access on its own — enforcement is a separate reviewed change.

  /** Mint a fresh CSPRNG token (~32 bytes, base64url). Exposed so a
   *  test / route can predict the shape; the store always generates
   *  its own in `createShareLink`. */
  static newShareToken(): string {
    return randomBytes(SHARE_TOKEN_BYTES).toString('base64url');
  }

  /** Insert a link token for `workbookId`. `password`, when given, is
   *  bcrypt-hashed with the same helper that protects user passwords —
   *  the plaintext never touches the DB. `expiresAt` is an absolute ms
   *  epoch or null for never. Returns the full persisted row. */
  createShareLink(opts: {
    workbookId: string;
    /** The collab room this token authorises. Bound at mint time so the
     *  enforcement gate can reject a token replayed against a different
     *  room (rooms are anonymous + not keyed by workbookId). */
    roomId: string;
    role: ShareRole;
    createdBy: number;
    expiresAt?: number | null;
    password?: string | null;
  }): ShareLink {
    const token = PersonalAuthStore.newShareToken();
    const now = Date.now();
    const expiresAt = opts.expiresAt ?? null;
    const passwordHash =
      opts.password != null && opts.password.length > 0
        ? bcrypt.hashSync(opts.password, BCRYPT_ROUNDS)
        : null;
    this.db
      .prepare(
        `INSERT INTO share_links
           (token, workbook_id, room_id, role, expires_at, password_hash, created_at, created_by)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
      )
      .run(
        token,
        opts.workbookId,
        opts.roomId,
        opts.role,
        expiresAt,
        passwordHash,
        now,
        opts.createdBy,
      );
    return {
      workbookId: opts.workbookId,
      roomId: opts.roomId,
      token,
      role: opts.role,
      expiresAt,
      passwordHash,
      createdAt: now,
      createdBy: opts.createdBy,
    };
  }

  /** All link tokens for a workbook, newest first. Expired rows are
   *  kept (sharing-model §8 q2 — history is preserved); callers that
   *  care about live access use `getLinkRole`. */
  listShareLinks(workbookId: string): ShareLink[] {
    const rows = this.db
      .prepare(
        `SELECT token, workbook_id, room_id, role, expires_at, password_hash, created_at, created_by
           FROM share_links WHERE workbook_id = ? ORDER BY created_at DESC`,
      )
      .all(workbookId) as ShareLinkRow[];
    return rows.map(rowToShareLink);
  }

  /** Read a single token, regardless of expiry. Returns null when
   *  unknown. */
  getShareLink(token: string): ShareLink | null {
    const row = this.db
      .prepare(
        `SELECT token, workbook_id, room_id, role, expires_at, password_hash, created_at, created_by
           FROM share_links WHERE token = ?`,
      )
      .get(token) as ShareLinkRow | undefined;
    return row ? rowToShareLink(row) : null;
  }

  /** Patch a token's role and/or expiry. Pass `expiresAt: null` to
   *  clear the expiry. Returns the updated row, or null if the token
   *  is unknown. */
  updateShareLink(
    token: string,
    patch: { role?: ShareRole; expiresAt?: number | null },
  ): ShareLink | null {
    const fields: string[] = [];
    const values: Array<unknown> = [];
    if (patch.role !== undefined) {
      fields.push('role = ?');
      values.push(patch.role);
    }
    if (patch.expiresAt !== undefined) {
      fields.push('expires_at = ?');
      values.push(patch.expiresAt);
    }
    if (fields.length > 0) {
      const result = this.db
        .prepare(`UPDATE share_links SET ${fields.join(', ')} WHERE token = ?`)
        .run(...values, token);
      if (result.changes === 0) return null;
    }
    return this.getShareLink(token);
  }

  /** Revoke a token. Returns false when the token was unknown. */
  deleteShareLink(token: string): boolean {
    const result = this.db.prepare('DELETE FROM share_links WHERE token = ?').run(token);
    return result.changes > 0;
  }

  /**
   * Resolve a token to its effective role, **respecting expiry** — an
   * expired token returns null as if it never existed.
   *
   * This is the read the future join-handshake will call to compute a
   * joiner's role. It is deliberately NOT called from `rooms.ts` yet:
   * tokens grant no access until that separate, reviewed batch wires
   * enforcement in. Surfaced + tested now so the persistence contract
   * is locked.
   */
  getLinkRole(token: string, now: number = Date.now()): ShareLinkRole | null {
    const link = this.getShareLink(token);
    if (!link) return null;
    if (link.expiresAt !== null && link.expiresAt <= now) return null;
    return {
      workbookId: link.workbookId,
      roomId: link.roomId,
      role: link.role,
      hasPassword: link.passwordHash !== null,
      passwordHash: link.passwordHash,
      expiresAt: link.expiresAt,
    };
  }

  // ── Internals ──────────────────────────────────────────────────────────

  private migrate(): void {
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE COLLATE NOCASE,
        password_hash TEXT NOT NULL,
        is_admin INTEGER NOT NULL DEFAULT 0,
        created_at INTEGER NOT NULL,
        display_name TEXT,
        email TEXT COLLATE NOCASE,
        timezone TEXT NOT NULL DEFAULT 'UTC',
        avatar_mime TEXT,
        avatar_blob BLOB,
        preferences TEXT NOT NULL DEFAULT '{}'
      );

      -- Email index — unique when present, allows NULL for users
      -- who didn't fill the field (the personal-mode docker doesn't
      -- need email since there's no SMTP recovery in v1).
      CREATE UNIQUE INDEX IF NOT EXISTS idx_users_email_unique
        ON users(email) WHERE email IS NOT NULL;

      CREATE TABLE IF NOT EXISTS sessions (
        id TEXT PRIMARY KEY,
        user_id INTEGER NOT NULL,
        expires_at INTEGER NOT NULL,
        FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
      );

      CREATE INDEX IF NOT EXISTS idx_sessions_user ON sessions(user_id);
      CREATE INDEX IF NOT EXISTS idx_sessions_expires ON sessions(expires_at);

      -- Files registry (Phase C Batch 2). Tracks which user owns each
      -- workbook stored in the HostIntegration backend. The byte
      -- payload lives there; we only carry metadata here so listing
      -- and ownership checks are fast (no per-file metadata read of
      -- the underlying object store).
      CREATE TABLE IF NOT EXISTS files (
        id TEXT PRIMARY KEY,
        owner_id INTEGER NOT NULL,
        display_name TEXT NOT NULL,
        size INTEGER NOT NULL,
        etag TEXT NOT NULL,
        created_at INTEGER NOT NULL,
        modified_at INTEGER NOT NULL,
        FOREIGN KEY (owner_id) REFERENCES users(id) ON DELETE CASCADE
      );

      CREATE INDEX IF NOT EXISTS idx_files_owner ON files(owner_id);
      CREATE INDEX IF NOT EXISTS idx_files_modified ON files(modified_at DESC);

      -- Share links (sharing-model §6.1). One row per minted token.
      -- The token is the PK + capability secret; password_hash is an
      -- optional layered bcrypt gate. These rows are persisted now but
      -- inert until the join-handshake batch reads getLinkRole — minting
      -- a token grants no access on its own. workbook_id is NOT a hard
      -- FK to files(id): tokens outlive the row only briefly, but more
      -- importantly the file registry is the ownership boundary checked
      -- at the route layer, not here.
      --
      -- room_id binds the token to a specific collab room (the Yjs
      -- documentName) at mint time. Rooms are anonymous and NOT keyed
      -- by workbook_id, so without this binding a token could be
      -- replayed against ANY room. The enforcement gate
      -- (resolveJoinRole) rejects a token whose room_id != the room
      -- being joined. Nullable at the SQL layer only so an ALTER on a
      -- pre-enforcement DB doesn't fail; new rows always populate it,
      -- and a null/empty room_id never matches a real room id (the gate
      -- rejects), so legacy inert rows stay inert.
      CREATE TABLE IF NOT EXISTS share_links (
        token TEXT PRIMARY KEY,
        workbook_id TEXT NOT NULL,
        room_id TEXT,
        role TEXT NOT NULL,
        expires_at INTEGER,
        password_hash TEXT,
        created_at INTEGER NOT NULL,
        created_by INTEGER NOT NULL,
        FOREIGN KEY (created_by) REFERENCES users(id) ON DELETE CASCADE
      );

      CREATE INDEX IF NOT EXISTS idx_share_links_workbook ON share_links(workbook_id);
    `);

    // Idempotent column migration — a pre-enforcement install created
    // share_links without room_id. Same guard pattern as the user
    // profile columns above (SQLite has no IF NOT EXISTS for ADD COLUMN).
    this.addColumnIfMissing('share_links', 'room_id', 'TEXT');

    // Idempotent column migrations — a Batch-1 install predates the
    // profile columns. Trying ALTER on every boot is cheap; the
    // INSTR check on PRAGMA table_info is the safe guard against
    // a duplicate-column error (SQLite has no IF NOT EXISTS for
    // ADD COLUMN).
    this.addColumnIfMissing('users', 'display_name', 'TEXT');
    this.addColumnIfMissing('users', 'email', 'TEXT COLLATE NOCASE');
    this.addColumnIfMissing('users', 'timezone', "TEXT NOT NULL DEFAULT 'UTC'");
    this.addColumnIfMissing('users', 'avatar_mime', 'TEXT');
    this.addColumnIfMissing('users', 'avatar_blob', 'BLOB');
    this.addColumnIfMissing('users', 'preferences', "TEXT NOT NULL DEFAULT '{}'");
    // The unique-email index is referenced by setEmail; safe-create it
    // here in case the migration block above didn't fire (CREATE TABLE
    // path already included it, but ALTER-upgraded DBs need it added).
    this.db.exec(`
      CREATE UNIQUE INDEX IF NOT EXISTS idx_users_email_unique
        ON users(email) WHERE email IS NOT NULL
    `);
  }

  private addColumnIfMissing(table: string, column: string, type: string): void {
    const cols = this.db.prepare(`PRAGMA table_info(${table})`).all() as Array<{ name: string }>;
    if (cols.some((c) => c.name === column)) return;
    this.db.exec(`ALTER TABLE ${table} ADD COLUMN ${column} ${type}`);
  }

  private applyBootstrap(spec: string): void {
    const [username, ...rest] = spec.split(':');
    const password = rest.join(':');
    if (!username || password.length < MIN_PASSWORD_LEN) return;
    if (!USERNAME_RE.test(username)) return;
    const hash = bcrypt.hashSync(password, BCRYPT_ROUNDS);
    this.db
      .prepare(
        'INSERT INTO users (username, password_hash, is_admin, created_at) VALUES (?, ?, 1, ?)',
      )
      .run(username, hash, Date.now());
  }
}

/** Raw `share_links` row shape as it comes back from SQLite. */
type ShareLinkRow = {
  token: string;
  workbook_id: string;
  /** Null only for legacy pre-enforcement rows (see the migration note).
   *  Normalised to '' by `rowToShareLink` so the gate's strict-equality
   *  room check treats it as "matches no room". */
  room_id: string | null;
  role: string;
  expires_at: number | null;
  password_hash: string | null;
  created_at: number;
  created_by: number;
};

function rowToShareLink(row: ShareLinkRow): ShareLink {
  return {
    workbookId: row.workbook_id,
    roomId: row.room_id ?? '',
    token: row.token,
    role: row.role as ShareRole,
    expiresAt: row.expires_at,
    passwordHash: row.password_hash,
    createdAt: row.created_at,
    createdBy: row.created_by,
  };
}

/** Parse `CASUAL_PERSONAL_MODE` to our enum, default `none`. */
export function readModeFromEnv(env: NodeJS.ProcessEnv = process.env): PersonalMode {
  const raw = (env.CASUAL_PERSONAL_MODE ?? 'none').toLowerCase();
  if (raw === 'single' || raw === 'multi' || raw === 'none') return raw;
  return 'none';
}
