/**
 * Embed playground — host side.
 *
 * Demonstrates the SDK's `<iframe>` delivery path end-to-end against the
 * SDK's own in-iframe runtime (`dist/embed/embed-runtime.js` + `embed.html`,
 * copied into ./public/embed/sheets/ at setup — see README).
 *
 * What the host does here:
 *   - Frames the SDK's embed.html and opens an `EmbedHostTransport`.
 *   - `onLoadRequest`  → serves the picked .xlsx bytes (or the bundled
 *                        sample) so the iframe can parse + mount.
 *   - `onSaveNotify`   → persists the snapshot JSON to localStorage + logs.
 *   - `onExit`         → same, on unmount.
 *   - `onError`        → logs the iframe's fatal boot/load errors.
 *   - buttons          → `sendCommandExecute` (bold/undo/redo) +
 *                        `sendSetViewMode` (editor ↔ preview).
 *
 * The editor never owns storage — it hands the host a snapshot and the
 * host decides. Here that's localStorage; a real host would PUT to WOPI,
 * Drive, etc.
 */

import { EmbedHostTransport } from '@casualoffice/sheets/embed';

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------

const DOC_ID = 'playground-doc';
const SNAPSHOT_KEY = `casual.embed.snapshot.${DOC_ID}`;
const SAMPLE_URL = './sample/sample.xlsx';
// The iframe + host are served from the same static origin, so the embed
// origin is just our own origin.
const EMBED_ORIGIN = window.location.origin;

// ---------------------------------------------------------------------------
// DOM
// ---------------------------------------------------------------------------

const iframe = document.getElementById('editor') as HTMLIFrameElement;
const fileInput = document.getElementById('file') as HTMLInputElement;
const saveBtn = document.getElementById('save') as HTMLButtonElement;
const viewModeSel = document.getElementById('view-mode') as HTMLSelectElement;
const boldBtn = document.getElementById('cmd-bold') as HTMLButtonElement;
const undoBtn = document.getElementById('cmd-undo') as HTMLButtonElement;
const redoBtn = document.getElementById('cmd-redo') as HTMLButtonElement;
const reloadBtn = document.getElementById('reload') as HTMLButtonElement;
const logEl = document.getElementById('log') as HTMLPreElement;

// ---------------------------------------------------------------------------
// Log helper — also mirrored to window for the Playwright spec to read.
// ---------------------------------------------------------------------------

type LogKind = 'info' | 'in' | 'out' | 'err';

interface LogEntry {
  ts: number;
  kind: LogKind;
  msg: string;
}

const logEntries: LogEntry[] = [];
// Exposed so tests (and curious devs) can assert on what crossed the wire.
(window as unknown as { __log: LogEntry[] }).__log = logEntries;

function log(kind: LogKind, msg: string): void {
  const entry: LogEntry = { ts: Date.now(), kind, msg };
  logEntries.push(entry);
  const time = new Date(entry.ts).toLocaleTimeString();
  const line = document.createElement('div');
  line.className = kind === 'info' ? '' : kind;
  line.innerHTML = `<span class="ts">${time}</span> `;
  line.appendChild(document.createTextNode(msg));
  logEl.appendChild(line);
  logEl.scrollTop = logEl.scrollHeight;
}

// ---------------------------------------------------------------------------
// The file the host will serve to the iframe on `onLoadRequest`. Starts as
// the bundled sample; a picked file replaces it.
// ---------------------------------------------------------------------------

let pendingBytes: ArrayBuffer | null = null;
let pendingName = 'sample.xlsx';

async function loadBytesForServe(): Promise<ArrayBuffer> {
  if (pendingBytes) return pendingBytes;
  log('info', `fetching bundled sample (${SAMPLE_URL})`);
  const res = await fetch(SAMPLE_URL);
  if (!res.ok) throw new Error(`sample fetch failed: ${res.status}`);
  return res.arrayBuffer();
}

fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0];
  if (!file) return;
  void file.arrayBuffer().then((buf) => {
    pendingBytes = buf;
    pendingName = file.name;
    log('info', `picked "${file.name}" (${buf.byteLength} bytes) — reloading editor`);
    mountIframe();
  });
});

// ---------------------------------------------------------------------------
// Transport lifecycle. Re-created each time the iframe (re)loads so the
// `iframeWindow` reference stays valid.
// ---------------------------------------------------------------------------

let transport: EmbedHostTransport | null = null;

function buildIframeSrc(): string {
  const viewMode = viewModeSel.value === 'preview' ? 'preview' : 'editor';
  // Pass-through the host's ?theme= so the embed can be exercised in dark mode.
  const theme = new URLSearchParams(window.location.search).get('theme') ?? 'system';
  const params = new URLSearchParams({ app: 'sheet', docId: DOC_ID, viewMode, theme });
  return `./embed/sheets/embed.html?${params.toString()}`;
}

function mountIframe(): void {
  transport?.destroy();
  transport = null;
  iframe.src = buildIframeSrc();
}

iframe.addEventListener('load', () => {
  const win = iframe.contentWindow;
  if (!win) {
    log('err', 'iframe has no contentWindow');
    return;
  }

  transport?.destroy();
  transport = new EmbedHostTransport({
    app: 'sheet',
    iframeWindow: win,
    embedOrigin: EMBED_ORIGIN,
  });

  transport.on({
    onEditorReady: (data) => {
      // The runtime sends `casual.ready` AND `casual.hello` at boot; the
      // EmbedTransport also replies to our hello with another ready. So
      // DON'T send hello from inside this handler or it ping-pongs
      // (ready → hello → ready → …). We send a single hello once below,
      // after wiring handlers.
      log('in', `editor ready — v${data.version ?? '?'} (${data.commit ?? '?'})`);
    },

    // Editor asks for bytes for `docId`. Serve the picked/sample file.
    onLoadRequest: async ({ docId }) => {
      log('in', `load.request docId=${docId}`);
      try {
        const bytes = await loadBytesForServe();
        log('out', `load.response ok — "${pendingName}" (${bytes.byteLength} bytes)`);
        return { ok: true, bytes, fileName: pendingName };
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        log('err', `load.response error — ${message}`);
        return { ok: false, code: 'load_failed', message };
      }
    },

    // Lightweight save notification (Ctrl/Cmd+S in the iframe, or our Save
    // button → casual.command.save). Carries the full snapshot JSON.
    onSaveNotify: ({ snapshot, reason }) => {
      persistSnapshot(snapshot);
      log('in', `save.notify (reason=${reason}) — snapshot persisted to localStorage`);
    },

    // Final snapshot on iframe unmount.
    onExit: ({ snapshot }) => {
      persistSnapshot(snapshot);
      log('in', 'exit — final snapshot persisted to localStorage');
    },

    onSelectionFormatState: (fmt) => {
      boldBtn.style.fontWeight = fmt.bold ? '700' : '400';
    },

    onError: ({ code, message }) => {
      log('err', `editor error [${code}] ${message}`);
    },
  });

  // Single host hello, advertising what this host can do. Fire-and-forget;
  // the runtime doesn't gate the load on it (it requests bytes on its own).
  transport.sendHostHello({ capabilities: ['load', 'save'] });
  log('info', `transport attached (${iframe.src})`);
});

function persistSnapshot(snapshot: unknown): void {
  try {
    localStorage.setItem(SNAPSHOT_KEY, JSON.stringify(snapshot));
    // Surface for the Playwright spec.
    (window as unknown as { __lastSnapshot: unknown }).__lastSnapshot = snapshot;
  } catch (err) {
    log('err', `persist failed — ${err instanceof Error ? err.message : String(err)}`);
  }
}

// ---------------------------------------------------------------------------
// Host UI → editor commands
// ---------------------------------------------------------------------------

saveBtn.addEventListener('click', () => {
  // Host-rendered Save → the runtime snapshots and emits save.notify(reason:host).
  transport?.sendCommandSave();
  log('out', 'command.save');
});

viewModeSel.addEventListener('change', () => {
  const viewMode = viewModeSel.value === 'preview' ? 'preview' : 'editor';
  transport?.sendSetViewMode({ viewMode });
  log('out', `set.viewmode ${viewMode}`);
});

boldBtn.addEventListener('click', () => {
  transport?.sendCommandExecute({ command: 'bold' });
  log('out', 'command.execute bold');
});

undoBtn.addEventListener('click', () => {
  transport?.sendCommandExecute({ command: 'undo' });
  log('out', 'command.execute undo');
});

redoBtn.addEventListener('click', () => {
  transport?.sendCommandExecute({ command: 'redo' });
  log('out', 'command.execute redo');
});

reloadBtn.addEventListener('click', () => {
  log('info', 'reloading editor');
  mountIframe();
});

// ---------------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------------

log('info', 'host booting — framing SDK embed.html');
mountIframe();
