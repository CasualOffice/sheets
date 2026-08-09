# Co-editing — design

Real-time collaborative editing layered on top of the single-user editor.
Available in the self-hosted Docker image; the GitHub Pages demo at `sheet.casualoffice.org` stays single-user.

---

## Goals

- Two browsers editing the same sheet see each other's edits within ≈250 ms.
- Anonymous sessions — anyone with the room URL can edit. No accounts.
- Password-protected rooms with role-based access (edit / view-only).
- In-memory only by default; optional Redis persistence (7-day TTL) for sessions that survive restarts.
- Single Docker image (`casualoffice/sheets`) — one command to self-host.

## Out of scope

- Persistence beyond room lifecycle (no Postgres, no S3, no WOPI).
- Auth / per-user accounts.
- Multi-room load balancing / horizontal scaling — single process, in-memory.

---

## Stack

| Concern          | Pick                                                |
| ---------------- | --------------------------------------------------- |
| Sync transport   | Yjs (CRDT) + Hocuspocus WebSocket server            |
| HTTP / WebSocket | Fastify + `@hocuspocus/server`                      |
| Persistence      | Redis (optional, 7-day TTL on Y.Doc binary updates) |
| Distribution     | Single multi-stage Dockerfile, Node 22 Alpine       |

---

## Architecture

```
┌────────────────────────── Browser ───────────────────────────┐
│                                                              │
│  Casual Sheets (built static bundle)                         │
│  ├── Univer OSS — grid + formulas + rendering                │
│  ├── Yjs ↔ Univer bridge (packages/sdk/src/collab/bridge.ts) │
│  │     subscribe → ICommandService.onMutationExecutedForCollab│
│  │     apply remote → executeCommand(…, { fromCollab: true })│
│  ├── CollabDriver.tsx — join/leave/reconnect state machine   │
│  ├── PresenceLayer.tsx — peer cursor overlay                 │
│  ├── AvatarStack.tsx — title-bar presence                    │
│  ├── HistoryPanel.tsx — per-room op log                      │
│  └── y-websocket-provider → wss://host/yjs                   │
│                                                              │
└────────────────────────────┬─────────────────────────────────┘
                             │ WebSocket /yjs
                             ▼
┌────────────────────── Node server ───────────────────────────┐
│                                                              │
│  Fastify (HTTP)                                              │
│  ├── GET  /                    serves the built web app      │
│  ├── GET  /r/:roomId           same SPA, room context        │
│  ├── POST /api/rooms           create room {password?, seed?}│
│  ├── GET  /api/rooms/:id/info  {needsPassword, hasSeed, …}  │
│  ├── POST /api/rooms/:id/seed  xlsx upload                   │
│  ├── GET  /api/rooms/:id/seed  download seed                 │
│  ├── POST /api/rooms/:id/snapshot  gzip snapshot upload      │
│  ├── GET  /api/rooms/:id/snapshot  joiner fast-path          │
│  └── GET  /health              liveness                      │
│                                                              │
│  Hocuspocus (WebSocket /yjs)                                 │
│  ├── Room registry Map<roomId, RoomState>                    │
│  ├── Password gate: SHA-256, close code 4401 on fail         │
│  ├── Op-log compaction on requestIdleCallback (Stage 6)      │
│  └── GC: throwaway rooms evicted after TTL; seeded/password  │
│          rooms kept indefinitely (or until Redis TTL expires) │
│                                                              │
└──────────────────────────────────────────────────────────────┘
```

---

## Yjs document schema

One `Y.Doc` per room. Univer state is transported as a versioned mutation log;
presence and app-owned side channels use separate Yjs maps:

```
Y.Doc
├─ Y.Array "ops"                 v2 mutation/snapshot records
├─ Y.Map "casual-charts"         app-owned chart models
└─ Y.Map "casual-comment-authors" comment author metadata
```

Mutation record:

```ts
{ v: 2, kind: 'op', c: clientId, s: sequence, t, id, p, u? }
```

`c:s` is stable identity. `s` increases within one Yjs client. Snapshot records
replace `id/p/u` with `{ wb, frontier }`, where `frontier[clientId]` is the
highest sequence already materialized in `wb`. The frontier means a concurrent
offline operation is replayed after the snapshot even if Yjs physically places
it before the snapshot record.

**What we don't sync:** computed formula results (`v` on a cell with `f`). Each client computes locally. Keeps payload small and avoids `RAND()` / `NOW()` divergence.

---

## Bridge contract

### Local edit → Yjs

1. Subscribe to `ICommandService.onMutationExecutedForCollab` — fires for `CommandType.MUTATION` only, including `syncOnly` mutations. See `vendor/univer/packages/core/src/services/command/command.service.ts:404`.
2. Allocate a stable `c:s` record and coalesce ordinary appends per microtask via
   `doc.transact`. Univer's `__splitChunk__` slices each retain a separate Yjs
   changeset so a large paste/copy cannot create one oversized frame.
3. Remember the exact locally-applied record id. It may be skipped only while
   it remains in the canonical applied prefix.

### Remote update → Univer

1. Observe `Y.Array "ops"` and derive the logical v2 plan by stable ids and the
   dominating snapshot frontier — never by array position.
2. If a concurrent insert changed the already-applied prefix, restore the
   preservation-aware initial/snapshot base and replay the final order,
   including this client's optimistic mutations.
3. `deepRewriteUnitId` patches the local workbook id in mutation params.
4. Await lazy-plugin registration and then await
   `executeCommand(id, params, { fromCollab: true })` before starting the next
   record. A `false` result is a failure.
5. Unknown, malformed, or failed records remain the unapplied head. Later
   records do not run and compaction is disabled. Hosts inspect
   `bridge.isReplayBlocked()` and the replay-failure subscription.

### Echo-loop prevention

Remote commands carry the bridge's private execution-options object (which also
contains `{ fromCollab: true }`). Echo suppression checks that object identity,
not the public boolean alone, so an ordinary host API caller cannot hide a
divergent local edit from the log. Local records are recognized by stable id;
after a base restore they are intentionally replayed.

## Protocol migration

Version 1 records had no sequence and were consumed with a positional
`appliedCount`. A concurrent Y.Array insertion before that count could be lost,
and the missing identity makes an in-place repair unknowable. V2 therefore
fails closed on any legacy record. Deployments must create a fresh room (or let
the room TTL expire) when upgrading active v1 rooms. Do not mix cached v1 and v2
clients in one room.

---

## Presence

Peer state is routed via **Yjs Awareness** (separate from the document — doesn't affect undo/redo):

```ts
provider.awareness.setLocalStateField('cursor', { sheetId, row, col });
provider.awareness.setLocalStateField('selection', { sheetId, range });
provider.awareness.setLocalStateField('liveEdit', { sheetId, row, col, value });
provider.awareness.setLocalStateField('user', { name, color, lastSeen });
```

`PresenceLayer.tsx` renders a `<canvas>` overlay that paints each peer's selection rect, cursor, and name label. Cursor positions are recomputed on scroll and on zoom changes so they stay pinned to the correct cell in frozen panes.

`AvatarStack.tsx` in the title bar shows up to 4 peer initials + a `+N` overflow chip. Tooltips show "Active now" or "Last seen Ns ago".

`LiveEditGhost.tsx` renders character-by-character preview in the peer's current edit cell.

---

## Security

- **Password gate**: `POST /api/rooms` accepts `{ password }`. Hashed with SHA-256 + constant-time compare. Failing the WS upgrade returns close code `4401`; the client routes this to a retry prompt.
- **View-only enforcement**: Hocuspocus tags the session role. On the client, `CollabDriver` sets `WorkbookEditablePermission = false` on the Univer workbook, blocking all mutations at the engine layer — not just the UI.
- **Known gap**: the server itself does not reject mutations from view-only WebSocket connections. A client that bypasses the Univer permission gate could still push ops. Server-side enforcement is tracked as a P0 for the next cycle.

---

## Room lifecycle

| Event                           | What happens                                                                |
| ------------------------------- | --------------------------------------------------------------------------- |
| `POST /api/rooms`               | New `Y.Doc`, optionally seeded from xlsx. Returns `{ roomId }`.             |
| WS connect `/yjs?room=X&p=<pw>` | Hocuspocus joins the Y.Doc; replays state to the joiner.                    |
| Last client disconnects         | Room marked idle; timer starts.                                             |
| Idle > `ROOM_TTL_MIN`           | Throwaway rooms (no password, no seed) evicted. Password/seeded rooms kept. |
| Redis configured                | Y.Doc binary updates persisted; survives server restart.                    |
| Redis TTL expires               | Room data purged after 7 days of inactivity.                                |

---

## Joiner fast-path

When the owner shares a room, the client uploads a gzipped `IWorkbookData` snapshot to `POST /api/rooms/:id/snapshot`. Joiners fetch it from `GET /api/rooms/:id/snapshot` (immutable-cached) and install it directly — skipping the xlsx parse entirely. Any ops that arrived after the snapshot was taken are replayed by the Yjs provider on connect.

---

## Op-log compaction (Stage 6)

Long-lived rooms accumulate records. Compaction defaults to `off`. A host that
explicitly selects `compaction: 'auto'` lets the lowest currently visible
awareness client periodically replace the known log with one snapshot record.
It can do so only when its applied ids exactly equal the current replay plan and
there is no pending, in-flight, failed, reordered, or unsupported local state.
The snapshot comes from `CasualSheetsAPI.getContent()`, never raw
`FWorkbook.save()`, so opaque xlsx and host resources survive. A null snapshot
is a no-op. Concurrent incomparable snapshots fail closed.

Use `compaction: 'manual'` to retain only the guarded
`bridge.forceCompact()` hook. Keep `compaction: 'off'` when a server owns
validated checkpoints; replay and awareness remain enabled, but browsers cannot
publish a room base. Awareness is not a partition-safe leader election system,
which is why automatic browser checkpoints require an explicit opt-in.

---

## Self-host

```sh
# Quick start — in-memory, no persistence:
docker run --rm -p 3000:3000 casualoffice/sheets:latest

# With Redis — rooms survive restarts:
docker compose up -d
```

See [`docs/DOCKERHUB.md`](./DOCKERHUB.md) for the full compose snippet and configuration reference.
