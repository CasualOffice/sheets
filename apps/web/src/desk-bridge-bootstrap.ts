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

// Surface unexpected runtime errors without the alarming full-width red
// monospace banner across the grid (which reads like a crash/dev artifact).
// Full detail goes to the console — DevTools is available in release builds —
// and the user sees a single compact, dismissible notice. Gated on desktop
// mode so plain web never installs these listeners.
if (typeof window !== 'undefined' && isDesktop()) {
  const notify = (detail: string) => {
    // eslint-disable-next-line no-console
    console.error('[deskApp]', detail);
    if (document.getElementById('__deskapp_err__')) return;
    const el = document.createElement('div');
    el.id = '__deskapp_err__';
    el.setAttribute('role', 'alert');
    el.style.cssText =
      'position:fixed;top:12px;right:12px;max-width:320px;z-index:99999;' +
      'display:flex;gap:10px;align-items:flex-start;' +
      'background:#fef2f2;color:#991b1b;border:1px solid #fecaca;border-radius:8px;' +
      'box-shadow:0 4px 14px rgba(0,0,0,0.12);padding:10px 12px;' +
      "font:13px/1.4 -apple-system,system-ui,'Segoe UI',sans-serif;";
    const text = document.createElement('div');
    text.style.flex = '1';
    text.textContent = 'Something went wrong. See the console for details.';
    const close = document.createElement('button');
    close.type = 'button';
    close.textContent = '✕';
    close.setAttribute('aria-label', 'Dismiss');
    close.style.cssText =
      'border:0;background:transparent;color:inherit;cursor:pointer;font-size:13px;line-height:1;padding:0;';
    close.addEventListener('click', () => el.remove());
    el.appendChild(text);
    el.appendChild(close);
    (document.body || document.documentElement).appendChild(el);
  };
  window.addEventListener('error', (e) => {
    notify(`[error] ${e.message} at ${e.filename}:${e.lineno}:${e.colno}`);
  });
  window.addEventListener('unhandledrejection', (e) => {
    const reason = e.reason as { message?: string } | undefined;
    notify(`[unhandled rejection] ${reason?.message ?? String(e.reason)}`);
  });
}

// Offline fonts (desktop only). The web build loads Inter + Material Symbols
// from the Google Fonts CDN (index.html); the Tauri app has no network, so we
// declare the same font families from locally-bundled woff2 instead.
//
// Inter weights are served by the desktop shell at `./fonts/` (relative to the
// editor's `--base` mount) and are NOT shipped in the web bundle.
//
// Material Symbols ships as a ~93 KB SUBSET (apps/web/public/fonts/
// material-symbols-outlined.subset.woff2) carrying only the ~150 icon ligatures
// this app actually renders, vs the ~3.8 MB full variable font. The subset
// lives in this repo's `public/` so it travels with the sheets dist and the
// desktop bundle picks it up automatically; regenerate + re-verify coverage via
// scripts/fonts/subset-material-symbols.py. The `local(...)` fallback uses a
// full OS-installed copy when present. All four variable axes (FILL/GRAD/opsz/
// wght) are kept so filled + weighted icons still shape.
if (typeof window !== 'undefined' && isDesktop() && !document.getElementById('__deskapp_fonts__')) {
  const css = `
@font-face{font-family:'Inter';font-style:normal;font-weight:400;font-display:swap;src:local('Inter Regular'),local('Inter-Regular'),url('./fonts/inter-400.woff2') format('woff2');}
@font-face{font-family:'Inter';font-style:normal;font-weight:500;font-display:swap;src:local('Inter Medium'),local('Inter-Medium'),url('./fonts/inter-500.woff2') format('woff2');}
@font-face{font-family:'Inter';font-style:normal;font-weight:600;font-display:swap;src:local('Inter SemiBold'),local('Inter-SemiBold'),url('./fonts/inter-600.woff2') format('woff2');}
@font-face{font-family:'Inter';font-style:normal;font-weight:700;font-display:swap;src:local('Inter Bold'),local('Inter-Bold'),url('./fonts/inter-700.woff2') format('woff2');}
@font-face{font-family:'Material Symbols Outlined';font-style:normal;font-weight:100 700;font-display:block;src:local('Material Symbols Outlined'),url('./fonts/material-symbols-outlined.subset.woff2') format('woff2');}`;
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
        save(bytes: ArrayBuffer, baselineSeq?: number): Promise<string | null>;
        saveAs(name: string, bytes: ArrayBuffer, baselineSeq?: number): Promise<string | null>;
        setDirty?(dirty: boolean): void;
        currentEditSeq?(): number;
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
     *  files. The Rust side writes chunks to a temp file and only swaps
     *  it into place on `commit_save_document` (atomic rename), so a
     *  half-written file never clobbers the original. Any chunk OR the
     *  commit throwing propagates so the editor reports a failed save. */
    async function chunkedWrite(path: string, buf: ArrayBuffer) {
      // Never atomically replace a good file with an empty one. A degenerate
      // or failed serialization that yielded 0 bytes would otherwise commit
      // over the original on disk — silent data loss. Throw so the caller
      // re-throws and the editor reports a failed save instead.
      if (buf.byteLength === 0) {
        throw new Error(`refusing to write an empty spreadsheet to ${path}`);
      }
      await inv('begin_save_document', { path });
      const view = new Uint8Array(buf);
      const CHUNK = 1 << 20;
      for (let offset = 0; offset < view.byteLength; offset += CHUNK) {
        const slice = view.subarray(offset, Math.min(offset + CHUNK, view.byteLength));
        await inv('write_save_chunk', { path, offset, bytes: Array.from(slice) });
      }
      // Atomic commit: swaps the temp file into the target path.
      await inv('commit_save_document', { path });
    }

    // Serialize writes so two overlapping saves — a fast double Ctrl+S, or a
    // Ctrl+S while a large save is still streaming chunks — can't interleave
    // their begin/write/commit IPC against the shared per-path temp file and
    // corrupt it. Each write waits for the previous to settle; the chain
    // survives a failed write (caught) so one rejected save doesn't wedge
    // every later one. The error still propagates to the caller that issued it.
    let writeChain: Promise<unknown> = Promise.resolve();
    function serializedWrite(path: string, buf: ArrayBuffer): Promise<void> {
      const run = writeChain.then(() => chunkedWrite(path, buf));
      writeChain = run.catch(() => undefined);
      return run;
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

    // Best-effort dirty tracking for the Rust close-guard. We keep the
    // current dirty state in a module-local boolean so we only fire the
    // transition (clean→dirty / dirty→clean) once and never spam IPC.
    // The Rust `set_window_dirty` command infers the window from the
    // caller. All calls are best-effort and must never throw.
    //
    // The `true` transition is driven by the editor itself via
    // `bridge.setDirty(true)` — App.tsx subscribes to the command bus
    // (`ICommandService.onMutationExecutedForCollab`, the project's only
    // sanctioned change hook) and calls it for every real mutation. That
    // catches toolbar edits, paste, fill and undo/redo that a raw DOM
    // keystroke heuristic misses (Univer's grid is a canvas with no
    // <input>). The `false` transition fires here on a successful save.
    let isDirty = false;
    // Monotonic edit counter — bumped on every edit signal from the editor
    // (App.tsx calls setDirty(true) from the Univer mutation hook), not just
    // the clean→dirty transition. A save snapshots it before writing and
    // re-checks after the commit, so an edit that lands mid-write keeps the
    // window dirty instead of being cleared (and silently lost on close).
    let editSeq = 0;
    function setWindowDirty(dirty: boolean) {
      if (dirty === isDirty) return;
      isDirty = dirty;
      try {
        void inv('set_window_dirty', { dirty }).catch(() => undefined);
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
      // Editor → bridge dirty signal. App.tsx calls this from the command-bus
      // mutation hook; save() clears it. Best-effort, never throws.
      setDirty(dirty: boolean) {
        // Bump on every edit signal (even while already dirty) so an in-flight
        // save can detect a change that landed during the write.
        if (dirty) editSeq++;
        setWindowDirty(dirty);
      },
      // Current edit counter. The save caller reads this at the instant it
      // serializes the workbook (synchronously, before the async encode +
      // IPC) and passes it back as `save(bytes, baselineSeq)`, so an edit that
      // landed *between serialization and the write* — not just during the
      // write — is detected and the window stays dirty. See save() below.
      currentEditSeq() {
        return editSeq;
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
        // A short read means the file shrank/changed under us mid-read (another
        // process truncated or replaced it — the same external edits the file
        // watcher reports). The buffer is sized to `total`, so the tail would be
        // zero-padded and parse as a baffling corruption; fail clearly instead
        // so the caller can re-open rather than open a silently-mangled file.
        if (offset < total) {
          throw new Error(
            `Only read ${offset} of ${total} bytes from ${path} — the file changed while ` +
              'opening. Try opening it again.',
          );
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
      async save(bytes: ArrayBuffer, baselineSeq?: number): Promise<string | null> {
        if (filePath) {
          // Reference point for "did the doc change since these bytes were
          // produced": the caller's edit counter at serialization time when
          // provided, else the counter now (covers only edits during the
          // write — the legacy behaviour).
          const seqAtStart = baselineSeq ?? editSeq;
          try {
            await serializedWrite(filePath, bytes);
          } catch (err) {
            console.error('[deskApp] save failed for', filePath, err);
            throw err;
          }
          // Only mark clean if no edit landed since the bytes were serialized;
          // otherwise the window would read "saved" with unsaved changes.
          if (editSeq === seqAtStart) setWindowDirty(false);
          return filePath;
        }
        return bridge!.saveAs('Untitled.xlsx', bytes, baselineSeq);
      },
      async saveAs(
        suggestedName: string,
        bytes: ArrayBuffer,
        baselineSeq?: number,
      ): Promise<string | null> {
        const newPath = (await inv('pick_save_path', { suggestedName })) as string | null;
        if (!newPath) return null;
        const seqAtStart = baselineSeq ?? editSeq;
        try {
          await serializedWrite(newPath, bytes);
        } catch (err) {
          console.error('[deskApp] saveAs failed for', newPath, err);
          throw err;
        }
        try {
          await inv('add_recent_file', { path: newPath });
        } catch {
          /* best-effort */
        }
        filePath = newPath;
        // Only mark clean if no edit landed while the write was in flight.
        if (editSeq === seqAtStart) setWindowDirty(false);
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

    // --- Theme plumbing --------------------------------------------------
    // The launcher passes its theme as `?theme=<system|light|dark>` and
    // emits a Tauri event `deskapp://theme` (payload `{ theme }`) when the
    // user flips it live. We resolve `system` against the OS colour scheme,
    // expose `themeMode` (raw) + `theme` (resolved 'light'/'dark') on the
    // bridge global, and re-broadcast as a DOM CustomEvent `deskapp:theme`
    // (detail `{ mode, resolved }`) which the `../theme` provider listens
    // for. Wrapped so a missing matchMedia / Tauri event API never throws
    // and the editor still boots. In iframe mode there's no Tauri event
    // bus, but matchMedia + the URL param still drive the resolved value.
    try {
      const themeBridge = bridge as unknown as {
        themeMode: 'system' | 'light' | 'dark';
        theme: 'light' | 'dark';
      };
      const parseMode = (): 'system' | 'light' | 'dark' => {
        const raw = url.searchParams.get('theme');
        return raw === 'light' || raw === 'dark' ? raw : 'system';
      };
      let themeMode = parseMode();
      const mq =
        typeof window.matchMedia === 'function'
          ? window.matchMedia('(prefers-color-scheme: dark)')
          : null;
      const resolve = (mode: 'system' | 'light' | 'dark'): 'light' | 'dark' =>
        mode === 'system' ? (mq?.matches ? 'dark' : 'light') : mode;

      const reapply = () => {
        const resolved = resolve(themeMode);
        themeBridge.themeMode = themeMode;
        themeBridge.theme = resolved;
        try {
          window.dispatchEvent(
            new CustomEvent('deskapp:theme', { detail: { mode: themeMode, resolved } }),
          );
        } catch {
          /* CustomEvent unsupported — best-effort */
        }
      };
      // Initial publish so the provider can read `window.__deskApp__.theme`
      // synchronously at module init and the event fires for late listeners.
      reapply();

      // OS scheme changes only matter while we're tracking `system`.
      if (mq) {
        const onMq = () => {
          if (themeMode === 'system') reapply();
        };
        if (typeof mq.addEventListener === 'function') mq.addEventListener('change', onMq);
        else if (typeof mq.addListener === 'function') mq.addListener(onMq);
      }

      // Live launcher theme changes arrive over the Tauri event bus.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const tauriEvent = (window as any).__TAURI__?.event;
      if (tauriEvent?.listen) {
        void tauriEvent
          .listen('deskapp://theme', (e: { payload?: { theme?: string } }) => {
            const next = e?.payload?.theme;
            if (next === 'system' || next === 'light' || next === 'dark') {
              themeMode = next;
              reapply();
            }
          })
          .catch(() => undefined);
      }
    } catch (err) {
      console.debug('[deskApp] theme plumbing failed', err);
    }

    // --- Cold-start boot overlay (top-level desktop only) ----------------
    // Univer's canvas can take 1–2s to initialise; without this the window
    // shows a blank/white flash (and white even in dark mode) until the
    // grid paints. We inject a full-window overlay synchronously here —
    // before React/Univer mount — themed from the resolved theme above so
    // it never flashes white in dark mode. App.tsx calls
    // `window.__deskApp__.dismissBoot()` once the workbook is ready; an
    // ~8s safety timer guarantees it can never stick. The overlay sits
    // above Univer's canvas and is fully removed on dismiss so it never
    // intercepts grid input.
    if (isTopLevel) {
      try {
        const themed = bridge as unknown as { theme?: 'light' | 'dark'; dismissBoot?: () => void };
        const dark = themed.theme === 'dark';
        const bg = dark ? '#1a1a1a' : '#ffffff';
        const fg = dark ? '#e6e6e6' : '#3c3c3c';
        const ring = dark ? '#3a3a3a' : '#e2e2e2';
        const accent = dark ? '#6aa3ff' : '#2563eb';
        const hasFile = !!filePath;
        const label = hasFile ? 'Opening…' : 'New spreadsheet…';

        if (!document.getElementById('__deskapp_boot_style__')) {
          const st = document.createElement('style');
          st.id = '__deskapp_boot_style__';
          st.textContent =
            // Below the 99999 error banner (so a boot-time error stays
            // visible) but above Univer's canvas chrome.
            '#__deskapp_boot__{position:fixed;inset:0;z-index:99998;display:flex;' +
            'flex-direction:column;align-items:center;justify-content:center;gap:16px;' +
            'opacity:1;transition:opacity .25s ease;font:14px/1.4 Inter,system-ui,' +
            'sans-serif;}' +
            '#__deskapp_boot__ .deskapp-boot__spinner{width:34px;height:34px;' +
            'border-radius:50%;border:3px solid var(--deskapp-boot-ring);' +
            'border-top-color:var(--deskapp-boot-accent);' +
            'animation:deskapp-boot-spin .8s linear infinite;}' +
            '@keyframes deskapp-boot-spin{to{transform:rotate(360deg);}}';
          (document.head || document.documentElement).appendChild(st);
        }

        const overlay = document.createElement('div');
        overlay.id = '__deskapp_boot__';
        overlay.style.background = bg;
        overlay.style.color = fg;
        overlay.style.setProperty('--deskapp-boot-ring', ring);
        overlay.style.setProperty('--deskapp-boot-accent', accent);
        // Brand mark — a simple grid glyph so we add no asset dependency
        // and stay safe before the editor's bundle/fonts have loaded.
        overlay.innerHTML =
          '<svg width="40" height="40" viewBox="0 0 24 24" fill="none" aria-hidden="true">' +
          '<rect x="3" y="3" width="18" height="18" rx="3" stroke="' +
          accent +
          '" stroke-width="2"/>' +
          '<path d="M3 9h18M3 15h18M9 3v18M15 3v18" stroke="' +
          accent +
          '" stroke-width="1.4" opacity="0.6"/></svg>' +
          '<div class="deskapp-boot__spinner"></div>' +
          '<div class="deskapp-boot__label">' +
          label +
          '</div>';
        (document.body || document.documentElement).appendChild(overlay);

        let dismissed = false;
        const dismissBoot = () => {
          if (dismissed) return;
          dismissed = true;
          try {
            const el = document.getElementById('__deskapp_boot__');
            if (!el) return;
            el.style.opacity = '0';
            el.style.pointerEvents = 'none';
            // Remove after the fade so it never intercepts grid input.
            window.setTimeout(() => el.remove(), 300);
          } catch {
            /* best-effort */
          }
        };
        themed.dismissBoot = dismissBoot;
        // Safety net: never let the overlay stick even if the ready signal
        // never fires (parse error swallowed, Univer stalls, etc.).
        window.setTimeout(dismissBoot, 8000);
      } catch (err) {
        console.debug('[deskApp] boot overlay failed', err);
      }
    }

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

    // External file-change listener (top-level windows only). The Rust
    // filesystem watcher emits `deskapp://file-changed` with payload
    // `{ kind: "modified"|"removed"|"renamed", path }` when the open
    // file is touched by another process. Re-broadcast as a DOM
    // CustomEvent so App.tsx can react without depending on __TAURI__.
    if (isTopLevel) {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const tauriEvent = (window as any).__TAURI__?.event;
        if (tauriEvent?.listen) {
          void tauriEvent
            .listen(
              'deskapp://file-changed',
              (e: { payload?: { kind?: string; path?: string } }) => {
                const { kind, path } = e?.payload ?? {};
                if (!kind || !path) return;
                try {
                  window.dispatchEvent(
                    new CustomEvent('deskapp:file-changed', { detail: { kind, path } }),
                  );
                } catch {
                  /* CustomEvent not supported */
                }
              },
            )
            .catch(() => undefined);
        }
      } catch {
        /* no Tauri event bus (web/iframe) — best-effort */
      }
    }
  }
}

declare global {
  interface Window {
    __deskApp__?: {
      isDesktop: true;
      filePath: string | null;
      loadDocument(p?: string): Promise<ArrayBuffer>;
      save(bytes: ArrayBuffer, baselineSeq?: number): Promise<string | null>;
      saveAs(name: string, bytes: ArrayBuffer, baselineSeq?: number): Promise<string | null>;
      /** Editor → bridge dirty signal for the Rust close-guard. Driven by
       *  App.tsx's command-bus mutation hook; cleared on save. */
      setDirty?(dirty: boolean): void;
      /** Edit counter read at serialization time and passed back to
       *  `save(bytes, baselineSeq)` so an edit between serialize and write
       *  keeps the window dirty. Top-level desktop bridge only. */
      currentEditSeq?(): number;
      /** Raw launcher theme preference: 'system' | 'light' | 'dark'. */
      themeMode?: 'system' | 'light' | 'dark';
      /** Resolved theme ('system' collapsed to 'light'/'dark'). */
      theme?: 'light' | 'dark';
      /** Idempotent: fade out + remove the cold-start boot overlay.
       *  Defined only in top-level desktop mode; safe to optional-chain. */
      dismissBoot?: () => void;
    };
  }
}

export {};
