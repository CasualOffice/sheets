/**
 * deskApp host bridge bootstrap — sheets edition.
 * Mirror of `docx/docx-editor/examples/vite/src/desk-bridge-bootstrap.ts`.
 * Keep the two in sync until we have a shared package.
 */

// Surface runtime errors visibly — the iframe context hides DevTools by
// default, so silent failures show up as a blank page. This overlay pins
// the first error to the top of the iframe so we can see it.
if (typeof window !== 'undefined') {
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

const url = new URL(window.location.href);
const isDesktop = url.searchParams.get('desk') === '1';
// eslint-disable-next-line no-console
console.log('[deskApp] bootstrap', { isDesktop, search: window.location.search });

if (isDesktop) {
  const isTopLevel = window.parent === window;
  let filePath = url.searchParams.get('file');

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
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
      get filePath() { return filePath; },
      // @ts-expect-error getter+setter on the same name
      set filePath(v: string | null) { filePath = v; },
      async loadDocument(p?: string): Promise<ArrayBuffer> {
        const path = p ?? filePath;
        if (!path) throw new Error('no file path bound to this window');
        return asArrayBuffer(await inv('load_document', { path }));
      },
      async save(bytes: ArrayBuffer): Promise<string | null> {
        if (filePath) {
          await inv('save_document', {
            path: filePath,
            bytes: Array.from(new Uint8Array(bytes)),
          });
          return filePath;
        }
        return bridge!.saveAs('Untitled.xlsx', bytes);
      },
      async saveAs(suggestedName: string, bytes: ArrayBuffer): Promise<string | null> {
        const written = (await inv('save_document_as', {
          suggestedName,
          bytes: Array.from(new Uint8Array(bytes)),
        })) as string | null;
        if (written) {
          filePath = written;
          await updateWindowTitleFromPath(written);
        }
        return written;
      },
    };
  } else {
    type RequestMethod = 'loadDocument' | 'save' | 'saveAs';
    let nextId = 0;
    const pending = new Map<number, { resolve: (v: unknown) => void; reject: (e: unknown) => void }>();

    function request<T>(method: RequestMethod, params: Record<string, unknown>): Promise<T> {
      return new Promise<T>((resolve, reject) => {
        const id = ++nextId;
        pending.set(id, {
          resolve: resolve as (v: unknown) => void,
          reject,
        });
        window.parent.postMessage(
          { src: 'deskApp', kind: 'request', id, method, params },
          '*',
        );
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
