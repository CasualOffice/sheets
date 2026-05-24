# Scaling

Single-process today; multi-process is on the v0.2 roadmap with
explicit caveats. Honest take below.

## Where the limits are (single process)

A 256 MiB container can comfortably host:

- ~100 concurrent rooms
- ~500 concurrent users (across rooms)
- ~50 MiB workbooks open
- ~1000 mutations/min sustained without GC pauses showing up

Real-world bottlenecks before CPU:

1. **Y.Doc memory pressure** — every active room holds its full
   CRDT graph in memory. Big workbooks with deep edit history
   inflate the doc; the Stage-6 compaction loop helps but doesn't
   help on workbooks held for a multi-day session. Plan ~1–2 MiB
   resident per active room at steady state.
2. **WebSocket throughput** — Node's default `ws` library handles
   ~10k concurrent connections on a modern x86 vCPU before event-
   loop latency creeps. Below that, the limit is fan-out of
   mutations to room peers, which is O(connections × mutations).
3. **xlsx I/O** — ExcelJS parse runs in Web Workers on the client
   side, so the server's only role in upload/download is bytes-in,
   bytes-out. No real bottleneck there until you saturate the
   network.

## Vertical scale recipes

- **Bigger RAM** — 1–2 GiB lets you host hundreds of rooms.
- **More CPU** — Node is single-threaded but `ws` does enough
  off-thread work that 2 vCPU helps the WebSocket I/O loop.
- **Increase `ROOM_TTL_MIN`** — rooms held in memory for 60 min
  after the last client leaves match the JIT working set of a
  workgroup using the app on a typical workday.

## Horizontal scale (v0.2 lane)

Multi-replica deployments work for **stateless paths** today —
File→Open / Download, WOPI, the admin REST — but the **WebSocket
collab plane has open items**:

1. **Sticky sessions** — clients in the same room must land on
   the same replica or they don't see each other. Use a load
   balancer's IP-hash policy on `/yjs` upgrades, OR pin via
   cookie. Caddy + nginx + Traefik all support this; the
   limitation is informational, not technical.
2. **Cross-replica awareness backplane** — Yjs awareness (peer
   cursors, presence) is in-memory per replica. Two users on the
   same room on different replicas see each other's mutations
   (via the Redis-persisted Y.Doc) but NOT each other's cursors.
   v0.2 ships a Redis pub/sub fan-out for awareness.
3. **Room creation race** — two clients hitting `POST /api/rooms`
   simultaneously on different replicas would create two rooms
   with the same id. Redis SETNX gate lands in v0.2.

So: today, multi-replica is fine for read-mostly fleets where
real-time collab is rare AND clients are pinned to a replica.
For a "real" multi-replica deployment with collab, wait for
v0.2.

## Operational signals to watch

```
GET /health        → { ok, rooms: <count> }
GET /api/rooms     → [{ id, clients, idleMs, ... }, ...]
GET /api/files/_health  → backend health probe
```

Plus the Fastify access log (`pino` JSON). Pipe into your
preferred log aggregator + alert on:

- `responseTime` p99 > 500 ms on `/wopi/files/:id/contents` —
  storage backend is slow.
- 5xx rate > 0.1% — something's wrong.
- WebSocket disconnects > expected — proxy timeout misconfig or
  the LB is killing long-lived connections.

## Self-host limits — when to upgrade infrastructure

| Signal | Likely cause | Move |
|---|---|---|
| `rooms` > 200 sustained | Healthy! | Add more RAM. |
| Container RSS > 80% of limit | Big-workbook hold | Bump memory; bump `ROOM_TTL_MIN` down to evict idle rooms sooner. |
| Frequent peer "Out of sync" pills | Network instability or proxy timeout | Bump proxy `read_timeout`. |
| Storage saves > 1 s p50 | Slow backend | Move `local` → `s3` or `postgres` co-located with the app. |

## Backups + DR

Covered in [`backups.md`](./backups.md). Two things outside the
backup story:

- The admin config JSON file holds secrets (`mode 0600`). It's
  part of `/data`; back it up.
- The `CASUAL_JWT_SECRET` env var is **not** in any backup. Lose
  it and every existing token becomes invalid. Restore from the
  deployment manifest / secret-manager.
