/**
 * HTTP-side load harness for the v0.1 capacity envelope.
 *
 * Drives the bounded WRITE endpoints (POST /api/rooms, POST
 * /seed, POST /snapshot, GET /snapshot) — i.e. exactly the surface
 * the rate-limit + room-cap work added in Streams C1/C2 protects.
 *
 * Output: a numbers table (count / errors / p50 / p95 / p99 ms per
 * endpoint) suitable for pasting into docs/LOAD_TEST.md.
 *
 * Usage:
 *   # Default: 50 virtual users, 60 s wall clock, target localhost:3000
 *   pnpm --filter @sheet/server load
 *
 *   # Override target + duration:
 *   LOAD_TARGET=http://localhost:3000 \
 *     LOAD_VUS=100 LOAD_DURATION_S=120 \
 *     pnpm --filter @sheet/server load
 *
 * Each virtual user runs a loop: create-room → seed (1 KB) →
 * snapshot upload (4 KB) → snapshot fetch. Latency is recorded per
 * step. VUs run in parallel; spin-up is staggered over 2 s so we
 * don't all hit /api/rooms at the same instant.
 *
 * Designed to be runnable WITHOUT installing k6 / artillery. Uses
 * only Node's built-in fetch + perf_hooks. Output is friendly to
 * being grepped or pasted into a doc.
 *
 * NOTE: This is HTTP-only. The Yjs WS sync path has its own scale
 * characteristics (more about per-room message fan-out than per-IP
 * request rate) and is a separate harness — TODO once we have a
 * provider available on the server side.
 */
import { performance } from 'node:perf_hooks';

const TARGET = process.env.LOAD_TARGET ?? 'http://localhost:3000';
const VUS = Number(process.env.LOAD_VUS ?? 50);
const DURATION_S = Number(process.env.LOAD_DURATION_S ?? 60);
const SPIN_UP_MS = Number(process.env.LOAD_SPIN_UP_MS ?? 2_000);

// 1 KB and 4 KB synthetic payloads — small enough that we measure
// the server-side bookkeeping (rate-limit middleware, multipart
// parse, registry write) rather than raw network throughput.
const SEED_BYTES = makeBytes(1024);
const SNAPSHOT_BYTES = makeBytes(4096);

interface StepMetrics {
  count: number;
  errors: number;
  rateLimited: number;
  latencies: number[];
}

const metrics: Record<string, StepMetrics> = {
  'POST /api/rooms': blankMetrics(),
  'POST /seed': blankMetrics(),
  'POST /snapshot': blankMetrics(),
  'GET /snapshot': blankMetrics(),
};

function blankMetrics(): StepMetrics {
  return { count: 0, errors: 0, rateLimited: 0, latencies: [] };
}

function makeBytes(n: number): Uint8Array {
  const out = new Uint8Array(n);
  for (let i = 0; i < n; i += 1) out[i] = i & 0xff;
  return out;
}

async function timed(
  step: keyof typeof metrics,
  fn: () => Promise<Response>,
): Promise<Response | null> {
  const m = metrics[step];
  const start = performance.now();
  try {
    const res = await fn();
    const dur = performance.now() - start;
    m.count += 1;
    m.latencies.push(dur);
    if (res.status === 429) m.rateLimited += 1;
    if (res.status >= 500) m.errors += 1;
    return res;
  } catch (err) {
    m.count += 1;
    m.errors += 1;
    m.latencies.push(performance.now() - start);
    return null;
  }
}

async function virtualUser(vuId: number, stopAt: number): Promise<void> {
  while (performance.now() < stopAt) {
    // 1. Create a fresh room.
    const create = await timed('POST /api/rooms', () =>
      fetch(`${TARGET}/api/rooms`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: '{}',
      }),
    );
    if (!create || !create.ok) {
      // Either the network blew up or we got rate-limited (429) —
      // back off briefly so we don't burn cycles re-hitting the
      // throttle.
      await sleep(250);
      continue;
    }
    const { roomId } = (await create.json()) as { roomId: string };

    // 2. Seed upload (small multipart). FormData on Node 20+ is
    //    native; no node-fetch dep needed.
    const seedForm = new FormData();
    seedForm.append(
      'file',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      new Blob([SEED_BYTES as any], { type: 'application/octet-stream' }),
      `vu${vuId}.bin`,
    );
    await timed('POST /seed', () =>
      fetch(`${TARGET}/api/rooms/${roomId}/seed`, {
        method: 'POST',
        body: seedForm,
      }),
    );

    // 3. Snapshot upload (raw gzip-ish bytes). Wrap in a Blob —
    //    fetch's BodyInit doesn't include Uint8Array in TS's lib
    //    types even though it works at runtime, and Blob is the
    //    sanctioned binary-body shape.
    await timed('POST /snapshot', () =>
      fetch(`${TARGET}/api/rooms/${roomId}/snapshot`, {
        method: 'POST',
        headers: { 'content-type': 'application/gzip' },
        body: // eslint-disable-next-line @typescript-eslint/no-explicit-any
new Blob([SNAPSHOT_BYTES as any]),
      }),
    );

    // 4. Snapshot fetch (the joiner fast-path). Read-only side.
    await timed('GET /snapshot', () => fetch(`${TARGET}/api/rooms/${roomId}/snapshot`));

    // Brief pause so a single VU doesn't pin one CPU at 100 % and
    // skew latency measurements vs. realistic Inter-arrival.
    await sleep(50 + Math.random() * 100);
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

function percentile(sorted: readonly number[], p: number): number {
  if (sorted.length === 0) return 0;
  // Nearest-rank for simplicity — small samples don't deserve
  // linear interpolation.
  const idx = Math.min(
    sorted.length - 1,
    Math.ceil((p / 100) * sorted.length) - 1,
  );
  return sorted[Math.max(0, idx)];
}

function formatRow(label: string, m: StepMetrics): string {
  const sorted = [...m.latencies].sort((a, b) => a - b);
  const p50 = percentile(sorted, 50).toFixed(1);
  const p95 = percentile(sorted, 95).toFixed(1);
  const p99 = percentile(sorted, 99).toFixed(1);
  return [
    label.padEnd(20),
    String(m.count).padStart(7),
    String(m.errors).padStart(6),
    String(m.rateLimited).padStart(8),
    p50.padStart(8),
    p95.padStart(8),
    p99.padStart(8),
  ].join(' ');
}

async function main(): Promise<void> {
  console.log(
    `[loadtest] target=${TARGET} vus=${VUS} duration=${DURATION_S}s spin-up=${SPIN_UP_MS}ms`,
  );

  // Quick health probe so a misconfigured TARGET fails fast.
  try {
    const h = await fetch(`${TARGET}/health`);
    if (!h.ok) {
      console.error(`[loadtest] /health returned ${h.status} — server up?`);
      process.exit(1);
    }
  } catch (err) {
    console.error(`[loadtest] /health unreachable at ${TARGET}: ${String(err)}`);
    process.exit(1);
  }

  const stopAt = performance.now() + DURATION_S * 1_000;
  const promises: Promise<void>[] = [];
  for (let i = 0; i < VUS; i += 1) {
    const delay = (i / VUS) * SPIN_UP_MS;
    promises.push(
      sleep(delay).then(() => virtualUser(i, stopAt)),
    );
  }
  await Promise.all(promises);

  // Pretty print.
  console.log('');
  console.log('endpoint              count  errors   429s  p50(ms)  p95(ms)  p99(ms)');
  console.log('-------------------- ------- ------ -------- -------- -------- --------');
  for (const [label, m] of Object.entries(metrics)) {
    console.log(formatRow(label, m));
  }
  console.log('');
  // Aggregate
  const total = Object.values(metrics).reduce((a, m) => a + m.count, 0);
  const errs = Object.values(metrics).reduce((a, m) => a + m.errors, 0);
  const rl = Object.values(metrics).reduce((a, m) => a + m.rateLimited, 0);
  console.log(
    `[loadtest] totals: ${total} requests, ${errs} errors (5xx), ${rl} rate-limited (429), ${(total / DURATION_S).toFixed(1)} req/s avg`,
  );
}

void main().catch((err) => {
  console.error('[loadtest] crashed:', err);
  process.exit(1);
});
