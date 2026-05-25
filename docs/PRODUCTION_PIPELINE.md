# Production-readiness pipeline

Tracks the work required to take Casual Sheets from "feature-complete
v0" to a credible v1.0 release. Five streams, ordered by leverage on
**user-trust failure modes** (a corrupted co-edit session destroys
more trust than a missing chart axis label).

The order is deliberate: reliability first (data divergence is the
worst failure), then type safety (prevents the next class of bugs),
then backend hardening (DoS resilience), then measurement (we don't
know our limits), then release.

| Stream | Title                                  | Status      |
|--------|----------------------------------------|-------------|
| A1     | Bridge replay retry + dead-letter      | in_progress |
| A2     | Surface dead-letter detail in indicator| pending     |
| B1     | Typed Univer facade wrapper            | pending     |
| C1     | Backend rate limit on uploads          | pending     |
| C2     | Request size + room eviction caps      | pending     |
| D1     | Load test script + baseline numbers    | pending     |
| E      | Tag v0.1 once A1/C1/C2/D1 are green    | pending     |

## Stream A — Co-edit reliability

The Yjs bridge is the highest-risk surface: every silent failure in
the replay loop is a silent divergence between peers. The audit
flagged "no retry, no dead-letter" as the single biggest gap.

### A1 — Replay retry with backoff (in_progress)

**Problem:** `cmdSvc.executeCommand(rec.id, params, { fromCollab: true })`
in the replay loop catches all errors with a single `.warn + tick
counter` path. Lazy-plugin chunk-load failures (transient — network
blip during webpack chunk fetch) get conflated with malformed-mutation
failures (permanent — bad params, unknown command id). The transient
class never gets a second chance and the room silently diverges.

**Approach:**

1. **Classify** errors at the catch site:
   - `ChunkLoadError` / "Loading chunk N failed" / "failed to fetch
     dynamically imported module" → **transient** (retry)
   - everything else → **permanent** (dead-letter immediately)
2. **Retry** transient failures with backoff: 300 ms, 900 ms, 2700 ms
   (three attempts). Out-of-order against later mutations is OK — if
   the original failed, peers might have failed too; recovery is
   best-effort.
3. **Dead-letter** ring buffer (cap 20) of `{ id, params, lastError,
   attempts, firstFailedAt, lastFailedAt, classification }`. Older
   entries evict on overflow. Exposed via `getReplayDeadLetter()` on
   the bridge handle.
4. **Keep** the `replayFailures` counter (existing UI consumes it).
   Only increment on **final** failure (after all retries exhausted
   for transient, immediately for permanent).

**Test plan:** unit-test the classifier + `withRetry` helper directly
(pure functions). E2E coverage via the existing replay tests is
unchanged.

### A2 — Dead-letter detail in `CollabIndicator` (pending)

CollabIndicator currently shows `"N not synced"` with no detail. Add a
click-to-expand panel showing the last 5 dead-letter entries: mutation
id + truncated error message + age. Lets the user (and us, in
production) self-diagnose what's actually failing instead of just
seeing a count.

## Stream B — Type safety

### B1 — Typed Univer facade wrapper (pending)

`grep -rn "as any" apps/web/src | wc -l` returns ~100 hits, almost
all at the FUniver → workbook / sheet / range boundary. Build
`apps/web/src/univer-facade.ts` with typed wrappers:

```ts
export function activeSheet(api: FUniver): TypedSheet | null;
export function setRangeValues(sheet: TypedSheet, ...): void;
```

Convert the highest-traffic callers first: `home-tab-actions.ts`,
`sheet-actions.ts`, `formula-refs.ts`, `flash-fill.ts`. Then add an
ESLint rule (`@typescript-eslint/no-explicit-any`) scoped to those
files to prevent regressions.

## Stream C — Backend hardening

### C1 — Rate limit on uploads (pending)

POST /api/docs is currently unauthenticated and unlimited. Token-
bucket per source IP: 10 uploads/minute, 50/hour. 429 + Retry-After on
overflow. Implementation: simple in-process LRU map keyed by IP. (S3-
backed shared limit lives behind the v1 host integration; out of
scope for now.)

### C2 — Request size + room eviction caps (pending)

Verify the upload body cap (Hocuspocus default may not apply to our
Express route). Add a per-process room cap (default 256) with LRU
eviction on idle (no clients for 30 min) so a sustained "upload + drop"
attack can't OOM the gateway.

## Stream D — Load + measurement

### D1 — k6 baseline (pending)

We don't know our scale ceiling. Script: ramp 1 → 100 concurrent
rooms × 2 clients/room sending one update every 2 s. Record p50/p99
broadcast latency, gateway RSS, RSS-per-room. Document numbers in
`docs/LOAD_TEST.md`. Target floor for v0.1: 100 rooms × 2 clients
without p99 > 500 ms or RSS > 1 GB.

## Stream E — v0.1 release

Cut once A1, C1, C2, D1 are merged. Includes `CHANGELOG.md`,
release notes, updated README badges. Pin known-good Univer commit
to prevent upstream churn during the release window.
