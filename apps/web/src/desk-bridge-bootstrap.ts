/**
 * deskApp host bridge bootstrap — sheets edition.
 * Mirror of `docx/docx-editor/examples/vite/src/desk-bridge-bootstrap.ts`.
 * Keep the two in sync until we have a shared package.
 *
 * Desktop mode is **opt-in and default-OFF**: it only activates when the
 * page is loaded with `?desk=1` in the URL, which the Casual Office Tauri
 * shell appends when it spawns the editor window. In a plain browser the
 * flag is absent, `isDesktop()` is false, and this module is a complete
 * no-op (no globals, no listeners, no behaviour change).
 */

/**
 * True only when running inside the Casual Office desktop shell, signalled
 * by the `?desk=1` URL flag the Tauri host appends. Cheap, pure, and safe
 * to call from any module — returns false in SSR/test contexts with no
 * `window`. This is the single source of truth for desktop detection.
 */
export function isDesktop(): boolean {
  if (typeof window === 'undefined') return false;
  try {
    return new URL(window.location.href).searchParams.get('desk') === '1';
  } catch {
    return false;
  }
}

// Surface runtime errors visibly — the iframe context hides DevTools by
// default, so silent failures show up as a blank page. This overlay pins
// the first error to the top of the iframe so we can see it. Gated on
// desktop mode so plain web never installs these listeners.
if (typeof window !== 'undefined' && isDesktop()) {
  const showError = (msg: string) => {
    if (document.getElementById('__deskapp_err__')) return;
    const div = document.createElement('div');
    div.id = '__deskapp_err__';
    div.style.cssText =
      'position:fixed;top:0;left:0;right:0;background:#dc2626;color:#fff;' +
      'padding:8px 12px;z-index:99999;font:12px/1.4 monospace;white-space:pre-wrap;';
    div.textContent = msg;
    (document.body || document.documentElement).appendChild(div);
  };
  window.addEventListener('error', (e) => {
    showError(`[error] ${e.message}\n  at ${e.filename}:${e.lineno}:${e.colno}`);
  });
  window.addEventListener('unhandledrejection', (e) => {
    showError(`[unhandled rejection] ${e.reason?.message ?? e.reason}`);
  });
}

// Offline fonts (desktop only). The web build loads Inter + Material Symbols
// from the Google Fonts CDN (index.html); the Tauri app has no network, so we
// declare the same font families from locally-bundled woff2 instead. Files are
// served by the desktop shell at `./fonts/` (relative to the editor's `--base`
// mount), NOT shipped in the web bundle — so this adds zero bytes to web. The
// `local(...)` fallbacks use an installed copy when present.
if (typeof window !== 'undefined' && isDesktop() && !document.getElementById('__deskapp_fonts__')) {
  const css = `
@font-face{font-family:'Inter';font-style:normal;font-weight:400;font-display:swap;src:local('Inter Regular'),local('Inter-Regular'),url('./fonts/inter-400.woff2') format('woff2');}
@font-face{font-family:'Inter';font-style:normal;font-weight:500;font-display:swap;src:local('Inter Medium'),local('Inter-Medium'),url('./fonts/inter-500.woff2') format('woff2');}
@font-face{font-family:'Inter';font-style:normal;font-weight:600;font-display:swap;src:local('Inter SemiBold'),local('Inter-SemiBold'),url('./fonts/inter-600.woff2') format('woff2');}
@font-face{font-family:'Inter';font-style:normal;font-weight:700;font-display:swap;src:local('Inter Bold'),local('Inter-Bold'),url('./fonts/inter-700.woff2') format('woff2');}
@font-face{font-family:'Material Symbols Outlined';font-style:normal;font-weight:100 700;font-display:block;src:local('Material Symbols Outlined'),url('./fonts/material-symbols-outlined.woff2') format('woff2');}`;
  const style = document.createElement('style');
  style.id = '__deskapp_fonts__';
  style.textContent = css;
  (document.head || document.documentElement).appendChild(style);
}

if (typeof window !== 'undefined' && isDesktop()) {
  const url = new URL(window.location.href);
  console.log('[deskApp] bootstrap', { isDesktop: true, search: window.location.search });

  const isTopLevel = window.parent === window;
  let filePath = url.searchParams.get('file');

  const tauriCore: { invoke?: (cmd: string, args?: unknown) => Promise<unknown> } | undefined =
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__TAURI__?.core;

  let bridge:
    | {
        isDesktop: true;
        filePath: string | null;
        loadDocument(p?: string): Promise<ArrayBuffer>;
        save(bytes: ArrayBuffer): Promise<string | null>;
        saveAs(name: string, bytes: ArrayBuffer): Promise<string | null>;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        getProfile?: () => Promise<any>;
      }
    | undefined;

  if (isTopLevel && tauriCore?.invoke) {
    const inv = tauriCore.invoke;
    // load_document returns tauri::ipc::Response on Rust side; binary
    // IPC means JS gets an ArrayBuffer directly. No JSON cost, no
    // truncation. save/saveAs still go through the JSON array path
    // until the Tauri 2 binary-input route is verified.
    const asArrayBuffer = (raw: unknown): ArrayBuffer => {
      if (raw instanceof ArrayBuffer) return raw;
      if (raw instanceof Uint8Array) {
        const u8 = raw;
        return u8.byteOffset === 0 && u8.byteLength === u8.buffer.byteLength
          ? (u8.buffer as ArrayBuffer)
          : (u8.slice().buffer as ArrayBuffer);
      }
      return new Uint8Array(raw as number[]).buffer as ArrayBuffer;
    };
    /** Chunked write — same motivation as loadDocument's chunked read.
     *  Avoids the JSON-number-array IPC truncation threshold for big
     *  files. */
    async function chunkedWrite(path: string, buf: ArrayBuffer) {
      await inv('begin_save_document', { path });
      const view = new Uint8Array(buf);
      const CHUNK = 1 << 20;
      for (let offset = 0; offset < view.byteLength; offset += CHUNK) {
        const slice = view.subarray(offset, Math.min(offset + CHUNK, view.byteLength));
        await inv('write_save_chunk', { path, offset, bytes: Array.from(slice) });
      }
    }

    async function updateWindowTitleFromPath(newPath: string) {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const w = (window as any).__TAURI__?.window;
        if (!w?.getCurrentWindow) return;
        const name = newPath.split(/[\\/]/).pop() || newPath;
        await w.getCurrentWindow().setTitle(`Spreadsheet — ${name}`);
      } catch {
        /* best-effort */
      }
    }
    bridge = {
      isDesktop: true,
      get filePath() {
        return filePath;
      },
      set filePath(v: string | null) {
        filePath = v;
      },
      async loadDocument(p?: string): Promise<ArrayBuffer> {
        const path = p ?? filePath;
        if (!path) throw new Error('no file path bound to this window');
        const lower = path.toLowerCase();
        // Chunked read in 1 MB slices to avoid IPC payload truncation.
        const total = (await inv('document_size', { path })) as number;
        const CHUNK = 1 << 20;
        const out = new Uint8Array(total);
        let offset = 0;
        while (offset < total) {
          const length = Math.min(CHUNK, total - offset);
          const chunk = asArrayBuffer(await inv('read_document_chunk', { path, offset, length }));
          out.set(new Uint8Array(chunk), offset);
          offset += chunk.byteLength;
          if (chunk.byteLength === 0) break;
        }
        // Magic-byte sniff for ZIP-based formats. .xlsx / .xlsm / .ods
        // are all renamed zips and must start with PK\003\004.
        // CSV/TSV/TAB are plain text and skip the check.
        const isZipFormat =
          lower.endsWith('.xlsx') || lower.endsWith('.xlsm') || lower.endsWith('.ods');
        if (isZipFormat) {
          const looksZip =
            out.byteLength >= 4 &&
            out[0] === 0x50 &&
            out[1] === 0x4b &&
            out[2] === 0x03 &&
            out[3] === 0x04;
          if (!looksZip) {
            const isOLE =
              out.byteLength >= 8 &&
              out[0] === 0xd0 &&
              out[1] === 0xcf &&
              out[2] === 0x11 &&
              out[3] === 0xe0;
            if (isOLE) {
              throw new Error(
                'This file is an OLE compound file (usually a password-protected ' +
                  'workbook or a legacy .xls format). Open it in Excel or LibreOffice ' +
                  'and Save As .xlsx (without a password), then try again.',
              );
            }
            throw new Error(
              "This file doesn't look like a valid spreadsheet. It's missing the ZIP " +
                'header expected for .xlsx/.xlsm/.ods. It may be corrupted or in another format.',
            );
          }
        }
        return out.buffer as ArrayBuffer;
      },
      async save(bytes: ArrayBuffer): Promise<string | null> {
        if (filePath) {
          await chunkedWrite(filePath, bytes);
          return filePath;
        }
        return bridge!.saveAs('Untitled.xlsx', bytes);
      },
      async saveAs(suggestedName: string, bytes: ArrayBuffer): Promise<string | null> {
        const newPath = (await inv('pick_save_path', { suggestedName })) as string | null;
        if (!newPath) return null;
        await chunkedWrite(newPath, bytes);
        try {
          await inv('add_recent_file', { path: newPath });
        } catch {
          /* best-effort */
        }
        filePath = newPath;
        await updateWindowTitleFromPath(newPath);
        return newPath;
      },
      // Profile exposed to the editor so it can show a local-user chip
      // in place of the collab Share button.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      async getProfile(): Promise<any> {
        return await inv('get_profile');
      },
    };
  } else {
    type RequestMethod = 'loadDocument' | 'save' | 'saveAs';
    let nextId = 0;
    const pending = new Map<
      number,
      { resolve: (v: unknown) => void; reject: (e: unknown) => void }
    >();

    function request<T>(method: RequestMethod, params: Record<string, unknown>): Promise<T> {
      return new Promise<T>((resolve, reject) => {
        const id = ++nextId;
        pending.set(id, {
          resolve: resolve as (v: unknown) => void,
          reject,
        });
        window.parent.postMessage({ src: 'deskApp', kind: 'request', id, method, params }, '*');
      });
    }

    window.addEventListener('message', (event) => {
      const data = event.data;
      if (!data || data.src !== 'deskApp' || data.kind !== 'reply') return;
      const pendingReq = pending.get(data.id);
      if (!pendingReq) return;
      pending.delete(data.id);
      if (data.error) pendingReq.reject(new Error(String(data.error)));
      else pendingReq.resolve(data.result);
    });

    bridge = {
      isDesktop: true,
      filePath,
      async loadDocument(p?: string): Promise<ArrayBuffer> {
        const bytes = await request<number[]>('loadDocument', { path: p ?? filePath });
        return new Uint8Array(bytes).buffer;
      },
      async save(bytes: ArrayBuffer): Promise<string | null> {
        const written = await request<string | null>('save', {
          bytes: Array.from(new Uint8Array(bytes)),
        });
        if (written) bridge!.filePath = written;
        return written;
      },
      async saveAs(suggestedName: string, bytes: ArrayBuffer): Promise<string | null> {
        const written = await request<string | null>('saveAs', {
          suggestedName,
          bytes: Array.from(new Uint8Array(bytes)),
        });
        if (written) bridge!.filePath = written;
        return written;
      },
    };
  }

  if (bridge) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__deskApp__ = bridge;

    // Ctrl/Cmd-H — focus the launcher window. Only fires in top-level
    // mode where __TAURI__.core.invoke is directly available.
    if (isTopLevel && tauriCore?.invoke) {
      const inv = tauriCore.invoke;
      window.addEventListener('keydown', (e) => {
        const meta = e.ctrlKey || e.metaKey;
        if (meta && !e.shiftKey && !e.altKey && e.key.toLowerCase() === 'h') {
          e.preventDefault();
          inv('focus_launcher_window').catch(() => undefined);
        }
      });
    }
  }
}

declare global {
  interface Window {
    __deskApp__?: {
      isDesktop: true;
      filePath: string | null;
      loadDocument(p?: string): Promise<ArrayBuffer>;
      save(bytes: ArrayBuffer): Promise<string | null>;
      saveAs(name: string, bytes: ArrayBuffer): Promise<string | null>;
    };
  }
}

export {};
