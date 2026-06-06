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

  // ── Internals ──────────────────────────────────────────────────────────

  private migrate(): void {
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE COLLATE NOCASE,
        password_hash TEXT NOT NULL,
        is_admin INTEGER NOT NULL DEFAULT 0,
        created_at INTEGER NOT NULL
      );

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
    `);
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

/** Parse `CASUAL_PERSONAL_MODE` to our enum, default `none`. */
export function readModeFromEnv(env: NodeJS.ProcessEnv = process.env): PersonalMode {
  const raw = (env.CASUAL_PERSONAL_MODE ?? 'none').toLowerCase();
  if (raw === 'single' || raw === 'multi' || raw === 'none') return raw;
  return 'none';
}
