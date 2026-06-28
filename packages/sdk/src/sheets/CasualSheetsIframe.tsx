/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * CasualSheetsIframe — the iframe-mounting variant of `<CasualSheets>`.
 * Sheet sibling of `@casualoffice/docs`'s `<CasualEditorIframe>`;
 * see doc 16 in the parent repo.
 *
 * Differs from CasualEditorIframe in one place: the embed URL params
 * carry `app=sheet` and the embed-runtime inside the iframe knows to
 * convert raw xlsx bytes into an `IWorkbookData` snapshot via
 * `xlsxToWorkbookData` before mounting `<CasualSheets>`.
 *
 * Public surface intentionally identical to CasualSheets. v0.6 will
 * rename CasualSheetsIframe → CasualSheets and the existing direct-
 * mount component → CasualSheetsDirect.
 */

import {
  forwardRef,
  useCallback,
  useEffect,
  useRef,
  type CSSProperties,
  type MutableRefObject,
} from 'react';

import { EmbedHostTransport } from '../embed/EmbedHostTransport';
import type {
  CasualErrorData,
  CommandExecuteData,
  LoadResponseData,
  SaveResponseData,
  SelectionChangedData,
  SelectionFormatStateData,
  TelemetryEventData,
} from '../embed/protocol';

/** What the host-side load/save handlers consume + return. The wrapper
 *  binds these to the host's FileSource via simple adapters. */
export interface HostFileBridge {
  open(docId: string): Promise<{ bytes: ArrayBuffer; name: string; etag?: string }>;
  save?(docId: string, bytes: ArrayBuffer, opts?: { etag?: string }): Promise<{ etag: string }>;
}

export interface CasualSheetsIframeRef {
  setViewMode(mode: 'preview' | 'editor'): void;
  iframe(): HTMLIFrameElement | null;
  /** Dispatch a formatting / navigation command (bold, italic, undo, …)
   *  against the iframe's active selection. `args` carries command payloads
   *  (e.g. font family/size, colour) — forwarded over the protocol. v0.6+. */
  executeCommand(command: CommandExecuteData['command'], args?: CommandExecuteData['args']): void;
}

export interface CasualSheetsIframeProps {
  /** Host-side bytes bridge. The wrapper round-trips load / save to the
   *  iframe through postMessage; bytes never live in the iframe's origin
   *  except in-memory while the workbook is open. */
  fileSource: HostFileBridge;
  docId: string;
  /** Default `editor`. Live changes push casual.command.set.viewmode. */
  viewMode?: 'preview' | 'editor';
  /** Default `/embed/sheets`. Consumer copies the SDK's
   *  `dist/embed/{embed.html, embed-runtime.js, embed-runtime.css}`
   *  to this path. */
  embedBasePath?: string;
  onSelectionChanged?: (data: SelectionChangedData) => void;
  /** Fires when the active cell's format flags change (bold, italic,
   *  …). Drive's custom toolbar reflects this state in the button
   *  "pressed" indicators. v0.6+. */
  onSelectionFormatState?: (data: SelectionFormatStateData) => void;
  onTelemetry?: (data: TelemetryEventData) => void;
  onError?: (data: CasualErrorData) => void;
  style?: CSSProperties;
  className?: string;
  testId?: string;
}

const DEFAULT_STYLE: CSSProperties = {
  width: '100%',
  height: '100%',
  border: 'none',
  display: 'block',
};

export const CasualSheetsIframe = forwardRef<CasualSheetsIframeRef, CasualSheetsIframeProps>(
  function CasualSheetsIframe(props, ref) {
    const {
      fileSource,
      docId,
      viewMode = 'editor',
      embedBasePath = '/embed/sheets',
      onSelectionChanged,
      onSelectionFormatState,
      onTelemetry,
      onError,
      style,
      className,
      testId = 'casual-sheets-iframe',
    } = props;

    const iframeRef = useRef<HTMLIFrameElement | null>(null);
    const transportRef = useRef<EmbedHostTransport | null>(null);
    const fileSourceRef = useRef(fileSource);
    fileSourceRef.current = fileSource;

    const onLoad = useCallback(async (req: { docId: string }): Promise<LoadResponseData> => {
      try {
        const { bytes, name, etag } = await fileSourceRef.current.open(req.docId);
        return {
          ok: true,
          bytes,
          fileName: name,
          ...(etag !== undefined ? { etag } : {}),
        };
      } catch (err) {
        return {
          ok: false,
          code: 'open_failed',
          message: err instanceof Error ? err.message : String(err),
        };
      }
    }, []);

    const onSave = useCallback(
      async (req: {
        docId: string;
        bytes: ArrayBuffer;
        baseEtag?: string;
      }): Promise<SaveResponseData> => {
        try {
          if (!fileSourceRef.current.save) {
            return {
              ok: false,
              code: 'save_unsupported',
              message: 'host fileSource does not implement save',
            };
          }
          const opts = req.baseEtag !== undefined ? { etag: req.baseEtag } : undefined;
          const { etag } = await fileSourceRef.current.save(req.docId, req.bytes, opts);
          return { ok: true, etag };
        } catch (err) {
          return {
            ok: false,
            code: 'save_failed',
            message: err instanceof Error ? err.message : String(err),
          };
        }
      },
      [],
    );

    const onIframeLoad = useCallback(() => {
      const iframe = iframeRef.current;
      if (!iframe?.contentWindow) return;
      transportRef.current?.destroy();
      const transport = new EmbedHostTransport({
        app: 'sheet',
        iframeWindow: iframe.contentWindow,
        embedOrigin: window.location.origin,
      });
      transport.on({
        onLoadRequest: onLoad,
        onSaveRequest: onSave,
        ...(onSelectionChanged ? { onSelectionChanged } : {}),
        ...(onSelectionFormatState ? { onSelectionFormatState } : {}),
        ...(onTelemetry ? { onTelemetry } : {}),
        ...(onError ? { onError } : {}),
        onEditorReady: () => {
          transport.sendHostHello({ capabilities: ['load', 'save'] });
          transport.sendSetViewMode({ viewMode });
        },
      });
      transportRef.current = transport;
    }, [
      onLoad,
      onSave,
      onSelectionChanged,
      onSelectionFormatState,
      onTelemetry,
      onError,
      viewMode,
    ]);

    useEffect(() => {
      transportRef.current?.sendSetViewMode({ viewMode });
    }, [viewMode]);

    useEffect(() => {
      return () => {
        transportRef.current?.destroy();
        transportRef.current = null;
      };
    }, []);

    if (ref) {
      const apiRef = ref as MutableRefObject<CasualSheetsIframeRef | null>;
      apiRef.current = {
        setViewMode: (mode) => transportRef.current?.sendSetViewMode({ viewMode: mode }),
        iframe: () => iframeRef.current,
        executeCommand: (command, args) =>
          transportRef.current?.sendCommandExecute({ command, args }),
      };
    }

    const url =
      `${embedBasePath}/embed.html` +
      `?app=sheet` +
      `&docId=${encodeURIComponent(docId)}` +
      `&viewMode=${viewMode}`;

    return (
      <iframe
        ref={iframeRef}
        src={url}
        onLoad={onIframeLoad}
        title="Casual Sheets"
        sandbox="allow-scripts allow-same-origin allow-downloads allow-modals"
        style={{ ...DEFAULT_STYLE, ...style }}
        className={className}
        data-testid={testId}
      />
    );
  },
);
