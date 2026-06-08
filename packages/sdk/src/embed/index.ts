/**
 * Embed — iframe delivery surface.
 *
 * EmbedTransport bridges postMessage to React handlers. The
 * envelope shapes (`./protocol.ts`) mirror the documentation
 * contract in `docs/internal/13-iframe-protocol.md` and the
 * signing payloads from `../signing/types.ts`.
 */

export {
  EmbedTransport,
  type EmbedTransportHandlers,
  type EmbedTransportOptions,
} from './EmbedTransport';
export {
  isCasualEnvelope,
  type CasualApp,
  type CasualEnvelope,
  type EditorHelloData,
  type HostHelloData,
  type LoadRequestData,
  type LoadResponseData,
  type LoadResponseDataOk,
  type LoadResponseDataErr,
  type SaveRequestData,
  type SaveResponseData,
  type SaveResponseDataOk,
  type SaveResponseDataErr,
  type SelectionChangedData,
  type TelemetryEventData,
  type LockLostData,
  type CommandSetReadOnlyData,
  type CommandSetThemeData,
  type CommandSetLocaleData,
  type SignatureRequestData,
  type SignatureRequestAckData,
  type SignatureFieldSignedData,
  type SignatureCompleteData,
  type SignatureCancelData,
} from './protocol';
