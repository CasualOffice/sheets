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
  EmbedHostTransport,
  type EmbedHostHandlers,
  type EmbedHostTransportOptions,
} from './EmbedHostTransport';
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
  type SaveNotifyData,
  type ExitData,
  type SelectionChangedData,
  type TelemetryEventData,
  type LockLostData,
  type CommandSetReadOnlyData,
  type CommandSetThemeData,
  type CommandSetLocaleData,
  type CommandSetFeaturesData,
  type DialogRequestData,
  type SignatureRequestData,
  type SignatureRequestAckData,
  type SignatureFieldSignedData,
  type SignatureCompleteData,
  type SignatureCancelData,
} from './protocol';
