/**
 * In-memory room registry. v1 — single process, no DB, no persistence
 * across restarts. See docs/CO-EDITING.md §"Room lifecycle".
 */

import { createHash, timingSafeEqual } from 'node:crypto';

type RoomState = {
  id: string;
  /** Number of currently connected websocket clients. */
  clients: number;
  /** Wall-clock ms when client count last hit zero. -1 while connected. */
  idleSince: number;
  /** Optional initial snapshot (e.g. from xlsx upload) — Hocuspocus seeds
   *  the Y.Doc with this on first connect. */
  seed?: Uint8Array;
  /** Optional xlsx bytes representing the room's *starting workbook*.
   *  Joiners fetch this once at /api/rooms/:id/seed and import it locally
   *  before the bridge takes over for incremental edits. The op-log alone
   *  doesn't carry pre-existing cells — the owner's "Share current
   *  workbook" flow uploads here so peers see the same starting state. */
  xlsxSeed?: Uint8Array;
  /** Optional gzipped JSON snapshot (`IWorkbookData`) — server-side cache
   *  so joiners skip the multi-second xlsx parse on join. Same content
   *  as `xlsxSeed`, just in a form the client can apply via
   *  `replaceWorkbook` without re-running ExcelJS. Lives alongside the
   *  xlsx (not instead of it) so existing tooling that wants the .xlsx
   *  bytes still has them. */
  snapshotGz?: Uint8Array;
  /** ISO timestamp the room was created. */
  createdAt: string;
  /** SHA-256 hash of the room password, hex-encoded. `null` = open room. */
  passwordHash: string | null;
};

const ROOM_TTL_MS = Number(process.env.ROOM_TTL_MIN ?? 5) * 60_000;
const GC_INTERVAL_MS = 30_000;

export class RoomRegistry {
  private rooms = new Map<string, RoomState>();
  private gc: ReturnType<typeof setInterval> | null = null;

  start(): void {
    if (this.gc) return;
    this.gc = setInterval(() => this.collect(), GC_INTERVAL_MS);
    this.gc.unref?.();
  }

  stop(): void {
    if (!this.gc) return;
    clearInterval(this.gc);
    this.gc = null;
  }

  /** Create a fresh room; returns its id. */
  create(opts: { password?: string; seed?: Uint8Array } = {}): string {
    const id = makeRoomId();
    this.rooms.set(id, {
      id,
      clients: 0,
      idleSince: Date.now(),
      seed: opts.seed,
      passwordHash: opts.password ? hashPassword(opts.password) : null,
      createdAt: new Date().toISOString(),
    });
    return id;
  }

  get(id: string): RoomState | undefined {
    return this.rooms.get(id);
  }

  /** Attach an xlsx-format starting workbook to a room. Returns false if
   *  the room doesn't exist. Idempotent — re-uploading replaces. */
  setXlsxSeed(id: string, bytes: Uint8Array): boolean {
    const room = this.rooms.get(id);
    if (!room) return false;
    room.xlsxSeed = bytes;
    return true;
  }

  /** Cache the gzipped JSON snapshot for fast-path joiner load. */
  setSnapshotGz(id: string, bytes: Uint8Array): boolean {
    const room = this.rooms.get(id);
    if (!room) return false;
    room.snapshotGz = bytes;
    return true;
  }

  /** Constant-time compare a user-supplied password against the room's hash. */
  passwordOk(id: string, candidate: string | undefined | null): boolean {
    const room = this.rooms.get(id);
    if (!room) return false;
    if (!room.passwordHash) return true; // open room
    if (!candidate) return false;
    const a = Buffer.from(hashPassword(candidate), 'hex');
    const b = Buffer.from(room.passwordHash, 'hex');
    if (a.length !== b.length) return false;
    return timingSafeEqual(a, b);
  }

  /** Idempotently mark a client joining the room. */
  onConnect(id: string): RoomState {
    const room = this.rooms.get(id) ?? {
      id,
      clients: 0,
      idleSince: Date.now(),
      createdAt: new Date().toISOString(),
      passwordHash: null,
    };
    room.clients += 1;
    room.idleSince = -1;
    this.rooms.set(id, room);
    return room;
  }

  onDisconnect(id: string): void {
    const room = this.rooms.get(id);
    if (!room) return;
    room.clients = Math.max(0, room.clients - 1);
    if (room.clients === 0) room.idleSince = Date.now();
  }

  /** For diagnostics — returns a snapshot, not a live reference. */
  snapshot() {
    return Array.from(this.rooms.values()).map((r) => ({
      id: r.id,
      clients: r.clients,
      idleSince: r.idleSince,
      createdAt: r.createdAt,
      protected: r.passwordHash !== null,
    }));
  }

  /** Evict rooms idle longer than ROOM_TTL_MS. */
  private collect(): void {
    const now = Date.now();
    for (const [id, room] of this.rooms) {
      if (room.idleSince > 0 && now - room.idleSince > ROOM_TTL_MS) {
        this.rooms.delete(id);
      }
    }
  }
}

/** Short room ids that are easy to share in a URL but unguessable enough
 *  for anonymous sessions. Base36, 12 chars ≈ 60 bits of entropy. */
function makeRoomId(): string {
  let out = '';
  for (let i = 0; i < 12; i++) {
    out += Math.floor(Math.random() * 36).toString(36);
  }
  return out;
}

/**
 * Hash a room password before storing. SHA-256 is enough — this is a
 * knowledge-of-secret gate, not an authentication-grade credential
 * (anonymous self-hosted v1 with no accounts). Salting would only help
 * if a memory dump and a rainbow-table attack were realistic threats;
 * neither is in scope.
 */
function hashPassword(plain: string): string {
  return createHash('sha256').update(plain, 'utf8').digest('hex');
}
