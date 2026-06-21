/**
 * EmbedTransport — postMessage bridge for the iframe-embedded
 * editor. Validates origin, dispatches envelopes to handlers,
 * sends responses back to the host. Pure TypeScript — no React.
 *
 * Wire shape defined in
 * `docs/internal/13-iframe-protocol.md` + `./protocol.ts`.
 *
 * Lifetime: constructed once when the /embed page mounts.
 * Destroyed when the page unloads. Concurrent calls into the same
 * handler are the host's responsibility (well-behaved hosts
 * serialise requests by id).
 */

import {
  isCasualEnvelope,
  type CasualApp,
  type CasualEnvelope,
  type EditorHelloData,
  type HostHelloData,
  type LoadResponseData,
  type LoadRequestData,
  type SaveRequestData,
  type SaveResponseData,
  type SaveNotifyData,
  type ExitData,
  type SelectionChangedData,
  type TelemetryEventData,
  type CommandSetReadOnlyData,
  type CommandSetThemeData,
  type CommandSetLocaleData,
  type CommandSetViewModeData,
  type CommandExecuteData,
  type SelectionFormatStateData,
  type CasualErrorData,
  type SignatureRequestData,
  type SignatureRequestAckData,
  type SignatureFieldSignedData,
  type SignatureCompleteData,
  type SignatureCancelData,
} from './protocol';

export interface EmbedTransportOptions {
  app: CasualApp;
  /** Origin allowed to send messages. Required — no `*` default. */
  hostOrigin: string;
  /** Editor build identity, surfaced in the editor.hello handshake. */
  version: string;
  commit: string;
  /** Editor-side capabilities advertised at handshake. */
  capabilities: string[];
  /** Optional injection for `window.parent` — tests pass a stub. */
  parentWindow?: {
    postMessage: (msg: unknown, targetOrigin: string, transfer?: Transferable[]) => void;
  };
  /** Optional injection for the window receiving messages — tests pass a stub. */
  hostWindow?: Pick<Window, 'addEventListener' | 'removeEventListener'>;
}

export interface EmbedTransportHandlers {
  /** Host → editor handshake. Editor responds with `editor.ready`. */
  onHostHello?: (data: HostHelloData) => void | Promise<void>;
  /** Host → editor: command.* messages. */
  onCommandSetReadOnly?: (data: CommandSetReadOnlyData) => void | Promise<void>;
  onCommandSetTheme?: (data: CommandSetThemeData) => void | Promise<void>;
  onCommandSetLocale?: (data: CommandSetLocaleData) => void | Promise<void>;
  /** Host → editor: switch chrome density (preview ↔ editor). */
  onCommandSetViewMode?: (data: CommandSetViewModeData) => void | Promise<void>;
  /** Host → editor: execute a formatting / navigation command (bold,
   *  italic, undo, …) against the active selection. v0.6+. */
  onCommandExecute?: (data: CommandExecuteData) => void | Promise<void>;
  onCommandFocus?: () => void | Promise<void>;
  onCommandSave?: () => void | Promise<void>;
  onCommandLoad?: () => void | Promise<void>;
  /** Host → editor signing session. Editor responds with `signature.request.ack`. */
  onSignatureRequest?: (
    data: SignatureRequestData,
  ) => SignatureRequestAckData | Promise<SignatureRequestAckData>;
  onSignatureCancel?: (data: SignatureCancelData) => void | Promise<void>;
  /** Host → editor: response to a prior load.request the editor issued. */
  onLoadResponse?: (id: string, data: LoadResponseData) => void;
  /** Host → editor: response to a prior save.request the editor issued. */
  onSaveResponse?: (id: string, data: SaveResponseData) => void;
}

export class EmbedTransport {
  private readonly opts: Required<Omit<EmbedTransportOptions, 'parentWindow' | 'hostWindow'>> & {
    parentWindow: NonNullable<EmbedTransportOptions['parentWindow']>;
    hostWindow: NonNullable<EmbedTransportOptions['hostWindow']>;
  };
  private handlers: EmbedTransportHandlers = {};
  private destroyed = false;
  private pendingRequests = new Map<string, (env: CasualEnvelope) => void>();

  constructor(opts: EmbedTransportOptions) {
    this.opts = {
      ...opts,
      parentWindow:
        opts.parentWindow ??
        (typeof window !== 'undefined'
          ? (window.parent as unknown as NonNullable<EmbedTransportOptions['parentWindow']>)
          : { postMessage: () => undefined }),
      hostWindow:
        opts.hostWindow ??
        (typeof window !== 'undefined'
          ? window
          : ({
              addEventListener: () => undefined,
              removeEventListener: () => undefined,
            } as unknown as NonNullable<EmbedTransportOptions['hostWindow']>)),
    };
    this.boundMessage = this.boundMessage.bind(this);
    this.opts.hostWindow.addEventListener('message', this.boundMessage as EventListener);
  }

  /** Replaces the handler set. Idempotent — call multiple times. */
  on(handlers: EmbedTransportHandlers): void {
    this.handlers = { ...this.handlers, ...handlers };
  }

  /** Editor → Host: announce ourselves. Call once after handlers are wired. */
  sendHello(): void {
    const data: EditorHelloData = {
      capabilities: this.opts.capabilities,
      version: this.opts.version,
      commit: this.opts.commit,
    };
    this.post('casual.hello', data);
  }

  /** Editor → Host: editor is ready to receive commands. */
  sendReady(): void {
    this.post('casual.ready', {});
  }

  /** Editor → Host: ask the host to supply document bytes for `docId`. */
  async requestLoad(docId: string, timeoutMs = 30000): Promise<LoadResponseData> {
    return this.request<LoadResponseData>(
      'casual.load.request',
      { docId } as LoadRequestData,
      timeoutMs,
    );
  }

  /** Editor → Host: ask the host to persist `bytes`. */
  async requestSave(req: SaveRequestData, timeoutMs = 30000): Promise<SaveResponseData> {
    return this.request<SaveResponseData>('casual.save.request', req, timeoutMs, [req.bytes]);
  }

  /** Editor → Host: the user asked to save (Ctrl/Cmd+S or a host save
   *  command). Fire-and-forget snapshot notification — the lightweight
   *  mirror of the React `onSave` hook. The host persists however it
   *  likes; no response is awaited. */
  sendSaveNotify(data: SaveNotifyData): void {
    this.post('casual.save.notify', data);
  }

  /** Editor → Host: the editor is unmounting. Carries the final snapshot
   *  so the host can persist on exit. Mirror of the React `onExit` hook. */
  sendExit(data: ExitData): void {
    this.post('casual.exit', data);
  }

  /** Editor → Host: selection moved. Fire-and-forget. */
  sendSelectionChanged(data: SelectionChangedData): void {
    this.post('casual.selection.changed', data);
  }

  /** Editor → Host: format flags on the active cell changed (bold,
   *  italic, …). Drive's custom toolbar mirrors this state in its
   *  "pressed" indicators. v0.6+. */
  sendSelectionFormatState(data: SelectionFormatStateData): void {
    this.post('casual.selection.format-state', data);
  }

  /** Editor → Host: a noteworthy event. */
  sendTelemetry(data: TelemetryEventData): void {
    this.post('casual.telemetry.event', data);
  }

  /** Editor → Host: per-field progress during a signing session. */
  sendSignatureFieldSigned(data: SignatureFieldSignedData): void {
    // Bytes ride the transfer list.
    this.post('casual.signature.field.signed', data, [data.bytes]);
  }

  /** Editor → Host: signing session is finished. */
  sendSignatureComplete(data: SignatureCompleteData): void {
    this.post('casual.signature.complete', data, [data.bytes]);
  }

  /** Editor → Host: editor-side cancel of a signing session. */
  sendSignatureCancel(reason: SignatureCancelData['reason']): void {
    this.post('casual.signature.cancel', { reason });
  }

  /** Editor → Host: fatal boot / load error. Host surfaces via wrapper
   *  `onError` callback. Fire-and-forget. */
  sendError(data: CasualErrorData): void {
    this.post('casual.error', data);
  }

  /** Tear down listeners. Idempotent. */
  destroy(): void {
    if (this.destroyed) return;
    this.opts.hostWindow.removeEventListener('message', this.boundMessage as EventListener);
    this.pendingRequests.clear();
    this.destroyed = true;
  }

  // ---------------------------------------------------------------
  // Internals
  // ---------------------------------------------------------------

  private boundMessage(ev: Event): void {
    // MessageEvent isn't always available on the typing side when
    // a stub window is injected; cast pragmatically.
    const msg = ev as unknown as MessageEvent;
    if (msg.origin && msg.origin !== this.opts.hostOrigin) return;
    if (!isCasualEnvelope(msg.data)) return;
    void this.dispatch(msg.data as CasualEnvelope);
  }

  private async dispatch(env: CasualEnvelope): Promise<void> {
    // Responses to outbound requests route through the correlation map.
    if (env.id && this.pendingRequests.has(env.id)) {
      const resolve = this.pendingRequests.get(env.id)!;
      this.pendingRequests.delete(env.id);
      resolve(env);
      return;
    }

    switch (env.type) {
      case 'casual.hello':
        await this.handlers.onHostHello?.(env.data as HostHelloData);
        this.sendReady();
        return;
      case 'casual.command.set.readonly':
        await this.handlers.onCommandSetReadOnly?.(env.data as CommandSetReadOnlyData);
        return;
      case 'casual.command.set.theme':
        await this.handlers.onCommandSetTheme?.(env.data as CommandSetThemeData);
        return;
      case 'casual.command.set.locale':
        await this.handlers.onCommandSetLocale?.(env.data as CommandSetLocaleData);
        return;
      case 'casual.command.set.viewmode':
        await this.handlers.onCommandSetViewMode?.(env.data as CommandSetViewModeData);
        return;
      case 'casual.command.execute':
        await this.handlers.onCommandExecute?.(env.data as CommandExecuteData);
        return;
      case 'casual.command.focus':
        await this.handlers.onCommandFocus?.();
        return;
      case 'casual.command.save':
        await this.handlers.onCommandSave?.();
        return;
      case 'casual.command.load':
        await this.handlers.onCommandLoad?.();
        return;
      case 'casual.signature.request': {
        const ack = (await this.handlers.onSignatureRequest?.(
          env.data as SignatureRequestData,
        )) ?? {
          ok: false,
          code: 'unhandled',
        };
        if (env.id) {
          this.postReply(env.id, 'casual.signature.request.ack', ack);
        }
        return;
      }
      case 'casual.signature.cancel':
        await this.handlers.onSignatureCancel?.(env.data as SignatureCancelData);
        return;
      default:
        // Unknown envelope — silently drop per the forward-compat
        // convention in the protocol doc.
        return;
    }
  }

  private post(type: string, data: unknown, transfer?: Transferable[]): void {
    const env: CasualEnvelope = { type, app: this.opts.app, v: 1, data };
    try {
      this.opts.parentWindow.postMessage(env, this.opts.hostOrigin, transfer);
    } catch {
      // Cross-origin postMessage can throw if the parent went away;
      // silently swallow — the next user action will retry.
    }
  }

  private postReply(id: string, type: string, data: unknown): void {
    const env: CasualEnvelope = { type, app: this.opts.app, id, v: 1, data };
    try {
      this.opts.parentWindow.postMessage(env, this.opts.hostOrigin);
    } catch {
      // see post()
    }
  }

  private async request<T>(
    type: string,
    data: unknown,
    timeoutMs: number,
    transfer?: Transferable[],
  ): Promise<T> {
    const id = newRequestId();
    return new Promise<T>((resolve, reject) => {
      const timer =
        timeoutMs > 0
          ? setTimeout(() => {
              this.pendingRequests.delete(id);
              reject(new Error(`Embed request ${type} timed out after ${timeoutMs}ms`));
            }, timeoutMs)
          : null;
      this.pendingRequests.set(id, (env) => {
        if (timer) clearTimeout(timer);
        resolve(env.data as T);
      });
      const env: CasualEnvelope = { type, app: this.opts.app, id, v: 1, data };
      try {
        this.opts.parentWindow.postMessage(env, this.opts.hostOrigin, transfer);
      } catch (err) {
        if (timer) clearTimeout(timer);
        this.pendingRequests.delete(id);
        reject(err);
      }
    });
  }
}

function newRequestId(): string {
  // 8 hex chars is plenty — IDs only need to be unique within a
  // single editor session. Math.random is fine; the host can't
  // use these for security.
  return Math.random().toString(16).slice(2, 10);
}
