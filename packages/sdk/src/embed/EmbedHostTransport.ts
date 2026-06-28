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
 * EmbedHostTransport — the parent side of the embed iframe bridge for
 * the sheet SDK. Mirror of @casualoffice/docs's host transport.
 *
 * Wire shape: `docs/SDK_SIGNING_EMBED.md` (cross-link in the doc repo at
 * `docs/internal/13-iframe-protocol.md`) + `16-sdk-iframe-architecture.md`.
 *
 * Lifetime: constructed once when the wrapper mounts the iframe.
 * `destroy()` removes the event listener; safe to call multiple times.
 */

import {
  isCasualEnvelope,
  type CasualApp,
  type CasualEnvelope,
  type CasualErrorData,
  type CommandSetReadOnlyData,
  type CommandSetThemeData,
  type CommandSetLocaleData,
  type CommandSetViewModeData,
  type CommandSetFeaturesData,
  type DialogRequestData,
  type CommandExecuteData,
  type SelectionFormatStateData,
  type EditorHelloData,
  type HostHelloData,
  type LoadRequestData,
  type LoadResponseData,
  type SaveRequestData,
  type SaveResponseData,
  type SaveNotifyData,
  type ExitData,
  type SelectionChangedData,
  type SignatureCancelData,
  type SignatureCompleteData,
  type SignatureFieldSignedData,
  type SignatureRequestData,
  type TelemetryEventData,
} from './protocol';

export interface EmbedHostTransportOptions {
  app: CasualApp;
  /** The iframe's `contentWindow`. */
  iframeWindow: Window;
  /** Origin allowed to send + receive messages. Same-origin internal
   *  embed is `window.location.origin`. */
  embedOrigin: string;
  /** Optional injection — tests pass a stub. */
  hostWindow?: Pick<Window, 'addEventListener' | 'removeEventListener'>;
}

export interface EmbedHostHandlers {
  onEditorReady?: (data: EditorHelloData) => void;
  /** Editor requests bytes for `docId`. */
  onLoadRequest?: (data: LoadRequestData) => Promise<LoadResponseData> | LoadResponseData;
  /** Editor requests a save (WOPI-style, carries xlsx bytes). */
  onSaveRequest?: (data: SaveRequestData) => Promise<SaveResponseData> | SaveResponseData;
  /** Editor fired its lightweight save notification (Ctrl/Cmd+S or a
   *  host save command). Carries the full snapshot JSON; fire-and-forget.
   *  Mirror of the React `onSave` hook. */
  onSaveNotify?: (data: SaveNotifyData) => void;
  /** Editor is unmounting; carries the final snapshot. Mirror of the
   *  React `onExit` hook. */
  onExit?: (data: ExitData) => void;
  onSelectionChanged?: (data: SelectionChangedData) => void;
  onSelectionFormatState?: (data: SelectionFormatStateData) => void;
  /** A host-owned dialog (Format Cells, Insert Chart, …) was requested from the
   *  editor's chrome. Render your own dialog and apply via executeCommand. */
  onDialogRequest?: (data: DialogRequestData) => void;
  onTelemetry?: (data: TelemetryEventData) => void;
  onSignatureFieldSigned?: (data: SignatureFieldSignedData) => void;
  onSignatureComplete?: (data: SignatureCompleteData) => void;
  onSignatureCancel?: (data: SignatureCancelData) => void;
  onError?: (data: CasualErrorData) => void;
}

type IframePostMessage = (msg: unknown, targetOrigin: string, transfer?: Transferable[]) => void;

export class EmbedHostTransport {
  private readonly opts: EmbedHostTransportOptions;
  private handlers: EmbedHostHandlers = {};
  private readonly boundOnMessage: (ev: MessageEvent) => void;
  private destroyed = false;

  constructor(opts: EmbedHostTransportOptions) {
    this.opts = opts;
    this.boundOnMessage = this.onMessage.bind(this);
    const target = opts.hostWindow ?? window;
    target.addEventListener('message', this.boundOnMessage);
  }

  on(handlers: EmbedHostHandlers): void {
    this.handlers = { ...this.handlers, ...handlers };
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    const target = this.opts.hostWindow ?? window;
    target.removeEventListener('message', this.boundOnMessage);
  }

  sendHostHello(data: HostHelloData): void {
    this.post('casual.hello', data);
  }

  sendSetViewMode(data: CommandSetViewModeData): void {
    this.post('casual.command.set.viewmode', data);
  }

  sendSetReadOnly(data: CommandSetReadOnlyData): void {
    this.post('casual.command.set.readonly', data);
  }

  sendSetTheme(data: CommandSetThemeData): void {
    this.post('casual.command.set.theme', data);
  }

  sendSetLocale(data: CommandSetLocaleData): void {
    this.post('casual.command.set.locale', data);
  }

  /** Host → editor: enable/disable chrome features (hide control + block command). */
  sendSetFeatures(data: CommandSetFeaturesData): void {
    this.post('casual.command.set.features', data);
  }

  sendCommandSave(): void {
    this.post('casual.command.save', null);
  }

  sendCommandFocus(): void {
    this.post('casual.command.focus', null);
  }

  /** Host → Editor: run a formatting / navigation command (bold,
   *  italic, undo, …) against the active selection. v0.6+. */
  sendCommandExecute(data: CommandExecuteData): void {
    this.post('casual.command.execute', data);
  }

  sendSignatureRequest(id: string, data: SignatureRequestData): void {
    this.post('casual.signature.request', data, id);
  }

  sendSignatureCancel(data: SignatureCancelData): void {
    this.post('casual.signature.cancel', data);
  }

  private onMessage(ev: MessageEvent): void {
    if (this.destroyed) return;
    if (ev.origin !== this.opts.embedOrigin) return;
    if (ev.source !== this.opts.iframeWindow) return;
    if (!isCasualEnvelope(ev.data)) return;
    if (ev.data.app !== this.opts.app) return;

    void this.dispatch(ev.data);
  }

  private async dispatch(env: CasualEnvelope): Promise<void> {
    switch (env.type) {
      case 'casual.ready':
        this.handlers.onEditorReady?.(env.data as EditorHelloData);
        return;
      case 'casual.load.request': {
        if (!this.handlers.onLoadRequest) return;
        const id = env.id ?? '';
        try {
          const resp = await this.handlers.onLoadRequest(env.data as LoadRequestData);
          const transfer: Transferable[] = resp.ok ? [resp.bytes] : [];
          this.post('casual.load.response', resp, id, transfer);
        } catch (err) {
          this.post(
            'casual.load.response',
            {
              ok: false as const,
              code: 'host_error',
              message: err instanceof Error ? err.message : String(err),
            },
            id,
          );
        }
        return;
      }
      case 'casual.save.request': {
        if (!this.handlers.onSaveRequest) return;
        const id = env.id ?? '';
        try {
          const resp = await this.handlers.onSaveRequest(env.data as SaveRequestData);
          this.post('casual.save.response', resp, id);
        } catch (err) {
          this.post(
            'casual.save.response',
            {
              ok: false as const,
              code: 'host_error',
              message: err instanceof Error ? err.message : String(err),
            },
            id,
          );
        }
        return;
      }
      case 'casual.save.notify':
        this.handlers.onSaveNotify?.(env.data as SaveNotifyData);
        return;
      case 'casual.exit':
        this.handlers.onExit?.(env.data as ExitData);
        return;
      case 'casual.selection.changed':
        this.handlers.onSelectionChanged?.(env.data as SelectionChangedData);
        return;
      case 'casual.selection.format-state':
        this.handlers.onSelectionFormatState?.(env.data as SelectionFormatStateData);
        return;
      case 'casual.dialog.request':
        this.handlers.onDialogRequest?.(env.data as DialogRequestData);
        return;
      case 'casual.telemetry.event':
        this.handlers.onTelemetry?.(env.data as TelemetryEventData);
        return;
      case 'casual.signature.field.signed':
        this.handlers.onSignatureFieldSigned?.(env.data as SignatureFieldSignedData);
        return;
      case 'casual.signature.complete':
        this.handlers.onSignatureComplete?.(env.data as SignatureCompleteData);
        return;
      case 'casual.signature.cancel':
        this.handlers.onSignatureCancel?.(env.data as SignatureCancelData);
        return;
      case 'casual.signature.request.ack':
        return;
      case 'casual.error':
        this.handlers.onError?.(env.data as CasualErrorData);
        return;
      default:
        return;
    }
  }

  private post(type: string, data: unknown, id?: string, transfer?: Transferable[]): void {
    const env: CasualEnvelope = {
      type,
      app: this.opts.app,
      v: 1,
      data,
      ...(id ? { id } : {}),
    };
    const send = this.opts.iframeWindow.postMessage.bind(
      this.opts.iframeWindow,
    ) as IframePostMessage;
    send(env, this.opts.embedOrigin, transfer);
  }
}
