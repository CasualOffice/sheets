/**
 * embed-runtime — the in-iframe entry point for the sheet SDK's iframe
 * delivery mode. Sheet sibling of `@schnsrw/docx-js-editor`'s
 * `embed-runtime`; see doc 16 in the parent repo.
 *
 * Lifecycle:
 *
 *   1. Wrapper renders <iframe src="${embedBasePath}/embed.html?...">.
 *   2. iframe loads embed.html which imports + runs `mountEmbedded()`
 *      from this module's compiled bundle.
 *   3. `mountEmbedded` parses URL params, opens an EmbedTransport,
 *      waits for the host's `casual.hello`, issues a load.request
 *      for the docId, parses the returned xlsx bytes via
 *      `xlsxToWorkbookData` → IWorkbookData, mounts <CasualSheets>
 *      with that snapshot into #casual-embed-root.
 *   4. Selection / autosave / signing events all bubble out via the
 *      same transport.
 */

import { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';

import { EmbedTransport } from '../embed/EmbedTransport';
import type { CasualApp } from '../embed/protocol';
import { CasualSheets } from '../sheets/CasualSheets';
import { xlsxToWorkbookData } from '../xlsx';
import type { IWorkbookData } from '@univerjs/core';

interface EmbedUrlConfig {
  app: CasualApp;
  docId: string;
  viewMode: 'preview' | 'editor';
}

function parseUrlConfig(search: string): EmbedUrlConfig {
  const params = new URLSearchParams(search);
  const app: CasualApp = params.get('app') === 'docs' ? 'docs' : 'sheet';
  const docId = params.get('docId') ?? '';
  const viewModeParam = params.get('viewMode');
  const viewMode: 'preview' | 'editor' = viewModeParam === 'editor' ? 'editor' : 'preview';
  return { app, docId, viewMode };
}

export interface MountEmbeddedOptions {
  root: HTMLElement;
  search?: string;
  hostOrigin?: string;
  identity?: { version: string; commit: string; capabilities?: string[] };
}

/** Public entry — called by embed.html. */
export function mountEmbedded(opts: MountEmbeddedOptions): void {
  const search = opts.search ?? (typeof window !== 'undefined' ? window.location.search : '');
  const config = parseUrlConfig(search);

  const hostOrigin =
    opts.hostOrigin ??
    inferHostOrigin() ??
    (typeof window !== 'undefined' ? window.location.origin : '');

  const identity = opts.identity ?? {
    version: '0.0.0',
    commit: 'unknown',
    capabilities: ['load', 'save', 'selection', 'signing'],
  };

  const transport = new EmbedTransport({
    app: config.app,
    hostOrigin,
    version: identity.version,
    commit: identity.commit,
    capabilities: identity.capabilities ?? ['load', 'save', 'selection', 'signing'],
  });

  opts.root.setAttribute('data-view-mode', config.viewMode);

  transport.on({
    onCommandSetViewMode: ({ viewMode }) => {
      opts.root.setAttribute('data-view-mode', viewMode);
      // v0.5.x will toggle CasualSheets's `ui` prop in response;
      // v0.5.0 only updates the data attribute so iframe-side CSS
      // can react.
    },
  });

  transport.sendHello();
  // Also send `casual.ready` so the host's `onEditorReady` fires without
  // waiting for the host to send `casual.hello` first. The protocol's
  // sendReady-on-receiving-hello path stays as a fallback for hosts that
  // do send hello eagerly; this just kicks off the handshake on the
  // iframe side so we don't deadlock when the host's strategy is to
  // wait for `casual.ready`.
  transport.sendReady();

  const reactRoot = createRoot(opts.root);
  reactRoot.render(<EmbeddedSheets transport={transport} docId={config.docId} />);
}

/** Inner React component — handles the async xlsx load + workbook
 *  mount. Splitting it out lets useEffect orchestrate the
 *  request/response without blocking the React tree at mount. */
function EmbeddedSheets({ transport, docId }: { transport: EmbedTransport; docId: string }) {
  const [data, setData] = useState<IWorkbookData | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const resp = await transport.requestLoad(docId);
        if (cancelled) return;
        if (!resp.ok) {
          setErrorMsg(resp.message ?? `load failed: ${resp.code}`);
          transport.sendError({
            code: 'load_failed',
            message: resp.message ?? resp.code,
          });
          return;
        }
        const snapshot = await xlsxToWorkbookData(resp.bytes);
        if (cancelled) return;
        setData(snapshot);
      } catch (err) {
        if (cancelled) return;
        const message = err instanceof Error ? err.message : String(err);
        setErrorMsg(message);
        transport.sendError({ code: 'parse_failed', message });
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [transport, docId]);

  if (errorMsg) {
    return (
      <div
        style={{
          width: '100%',
          height: '100%',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          padding: 24,
          textAlign: 'center',
          color: 'var(--danger, #d63a2f)',
          fontSize: 13,
        }}
      >
        Failed to load workbook: {errorMsg}
      </div>
    );
  }

  if (!data) {
    return (
      <div
        style={{
          width: '100%',
          height: '100%',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          color: 'var(--text-muted, #5a5a5a)',
          fontSize: 13,
        }}
      >
        Loading workbook…
      </div>
    );
  }

  return <CasualSheets initialData={data} />;
}

function inferHostOrigin(): string | undefined {
  if (typeof document === 'undefined') return undefined;
  const ref = document.referrer;
  if (!ref) return undefined;
  try {
    return new URL(ref).origin;
  } catch {
    return undefined;
  }
}
