# Load test — v0.1 baseline

HTTP-side capacity baseline for the v0.1 release. Measured with the
in-tree harness at `apps/server/scripts/loadtest.ts`. Drives the four
write-path endpoints the audit identified as the abuse surface: room
create, seed upload, snapshot upload, snapshot fetch.

Re-run with:

```bash
pnpm --filter @sheet/server load
# or:
LOAD_TARGET=http://localhost:3000 LOAD_VUS=50 LOAD_DURATION_S=20 \
  pnpm --filter @sheet/server load
```

## Test environment

| Variable | Value |
|---|---|
| Host | MacBook Air (Apple Silicon), darwin 25 |
| Node | 22.x (pnpm-managed) |
| Server | apps/server (Fastify 5 + @hocuspocus/server 2.15) |
| Storage | in-memory (no Redis attached) |
| Workbook host | memory |
| Rate-limit plugin | @fastify/rate-limit 10.3.0 |
| Room cap | MAX_ROOMS = 10 000 for baseline (raised so capacity wasn't the bottleneck) |
| Payloads | 1 KB seed, 4 KB snapshot — chosen to measure server bookkeeping, not network throughput |

## Run 1 — Baseline throughput (rate-limit DISABLED)

50 virtual users for 20 s; spin-up staggered over 2 s.
`RATE_LIMIT_ENABLED=false` so the bucket doesn't cap the result.

```
endpoint              count  errors   429s  p50(ms)  p95(ms)  p99(ms)
-------------------- ------- ------ -------- -------- -------- --------
POST /api/rooms         9519      0        0      0.2      0.6      1.6
POST /seed              9519      0        0      0.2      0.6      2.1
POST /snapshot          9519      0        0      0.2      0.5      1.9
GET /snapshot           9519      0        0      0.2      0.5      1.7

totals: 38 076 requests, 0 errors, 0 rate-limited, 1903.8 req/s avg
```

**Reading:** all four endpoints sustain ~1900 req/s combined (~480
req/s each) with p99 well under 3 ms. Zero 5xx, zero memory pressure
(in-memory host, no Redis I/O). The server is comfortably faster than
the rate-limit bucket needs to be — the bucket is the safety net,
not the bottleneck.

## Run 2 — Rate-limit verification (defaults ON)

20 virtual users for 15 s. Default `RATE_LIMIT_PER_MIN=60`,
`UPLOAD_RATE_LIMIT_PER_MIN=12`. Single source IP (the loadtest
harness) — designed to confirm the bucket triggers, not measure
real throughput.

```
endpoint              count  errors   429s  p50(ms)  p95(ms)  p99(ms)
-------------------- ------- ------ -------- -------- -------- --------
POST /api/rooms         1162      0     1102      0.9      1.7      2.8
POST /seed                60      0       48      0.6      1.6      3.7
POST /snapshot            60      0       48      0.4      0.9      2.6
GET /snapshot             60      0        0      0.3      0.7      1.6

totals: 1342 requests, 0 errors, 1198 rate-limited (89.3% throttled), 89.5 req/s avg
```

**Reading:** the bucket cuts the harness's offered load to exactly
the configured envelope.
- `/api/rooms`: 1162 attempts → 60 accepted (rest 429). Matches
  `RATE_LIMIT_PER_MIN=60` for a single IP across the 15 s window
  (Fastify's bucket allows the initial burst then enforces the rate).
- `/seed` and `/snapshot`: 60 attempts → 12 accepted. Matches
  `UPLOAD_RATE_LIMIT_PER_MIN=12`.
- `GET /snapshot`: 60 attempts → 0 throttled. The read endpoint
  is correctly NOT rate-limited (returning peers shouldn't get
  throttled for re-joining).

Zero 5xx in both runs — the rate-limit middleware is the only
pushback, exactly as designed.

## v0.1 SLO floor

Based on these numbers, the v0.1 floor (single-process, in-memory
host, no Redis):

- **HTTP write capacity per IP:** 60 room-creates + 12 uploads per
  minute (configurable via env).
- **HTTP write capacity per process:** > 1500 req/s aggregate
  (well above any realistic legitimate workload at v0.1 scale).
- **Latency:** p99 < 5 ms for all four write endpoints under the
  baseline run. The rate-limit middleware adds < 1 ms p99 overhead.
- **Concurrent rooms:** MAX_ROOMS = 256 default. When at cap, oldest
  evictable room is dropped (see Stream C2 in PRODUCTION_PIPELINE.md);
  if every slot is non-evictable, returns 503 + `retry-after: 60`.

## Out of scope (follow-up)

- **WS sync capacity.** The HTTP harness doesn't drive the
  `/yjs` WebSocket path. The audit-recommended floor (100 rooms ×
  2 clients × 1 update / 2 s without p99 > 500 ms) needs a Yjs
  provider-based harness — separate effort, won't block v0.1.
- **Redis-backed runs.** Numbers above are in-memory only. Redis
  adds 0.5–2 ms per persisted update; should re-run when Redis is
  the configured storage backend.
- **Multi-IP load.** Single-process, single-source-IP run. A real
  abuse pattern from 100 distinct IPs is bounded by `MAX_ROOMS`
  and per-IP buckets but the aggregate throughput cap is the
  server's CPU, not the rate-limit — that's the next thing to
  measure once D1 ships.

## Re-running

```bash
# 1. Start the server (default env)
pnpm --filter @sheet/server dev

# 2. In another shell:
pnpm --filter @sheet/server load              # 50 VUs × 60 s
LOAD_VUS=100 LOAD_DURATION_S=120 \
  pnpm --filter @sheet/server load            # 100 VUs × 2 min

# Or for raw-capacity numbers (no bucket in the way):
RATE_LIMIT_ENABLED=false MAX_ROOMS=10000 \
  pnpm --filter @sheet/server dev
# then run the loadtest as above.
```

The harness uses Node's built-in `fetch` + `perf_hooks` — no
k6 / artillery install needed. Output is grep-friendly so CI can
extract the p99 numbers later if we want a regression gate.
