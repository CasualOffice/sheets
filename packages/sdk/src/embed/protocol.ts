/**
 * Iframe protocol — wire envelopes that match
 * `docs/internal/13-iframe-protocol.md`.
 *
 * Mirror updates to the doc whenever a new envelope shape lands.
 * The discriminator on `type` (always starts with `casual.`) and
 * the per-envelope `data` shape are the contract.
 */

import type {
  CancelReason,
  SignatureCompletePayload,
  SignatureField,
  SignatureMode,
  SignedFieldPayload,
} from '../signing/types';

export type CasualApp = 'docs' | 'sheet';

/** Common envelope shape — every postMessage on the wire matches this. */
export interface CasualEnvelope<T = unknown> {
  type: string;
  app: CasualApp;
  /** Per-request id for request/response correlation. Empty for fire-and-forget. */
  id?: string;
  /** Protocol version. Bumped only on breaking changes. */
  v: 1;
  data: T;
}

// ---------------------------------------------------------------
// Handshake
// ---------------------------------------------------------------

export interface EditorHelloData {
  capabilities: string[];
  version: string;
  commit: string;
}

export interface HostHelloData {
  capabilities: string[];
  authToken?: string;
}

// ---------------------------------------------------------------
// Load + save (editor → host requests; host → editor responses)
// ---------------------------------------------------------------

export interface LoadRequestData {
  docId: string;
}

export interface LoadResponseDataOk {
  ok: true;
  bytes: ArrayBuffer;
  etag?: string;
  fileName: string;
  readOnly?: boolean;
}

export interface LoadResponseDataErr {
  ok: false;
  code: string;
  message?: string;
}

export type LoadResponseData = LoadResponseDataOk | LoadResponseDataErr;

export interface SaveRequestData {
  docId: string;
  bytes: ArrayBuffer;
  baseEtag?: string;
}

export interface SaveResponseDataOk {
  ok: true;
  etag: string;
}

export interface SaveResponseDataErr {
  ok: false;
  code: string;
  etag?: string;
  message?: string;
}

export type SaveResponseData = SaveResponseDataOk | SaveResponseDataErr;

// ---------------------------------------------------------------
// Selection + telemetry + lock (editor → host notifications)
// ---------------------------------------------------------------

export interface SelectionChangedData {
  docs?: { paraId: string; from: number; to: number; selectedText: string };
  sheet?: { sheet: string; from: string; to: string };
}

export interface TelemetryEventData {
  kind: string;
  /** Arbitrary metric fields. */
  [k: string]: unknown;
}

export interface LockLostData {
  reason: 'taken_by_other' | 'expired' | 'host_revoked';
}

// ---------------------------------------------------------------
// Commands (host → editor)
// ---------------------------------------------------------------

export interface CommandSetReadOnlyData {
  readOnly: boolean;
}

export interface CommandSetThemeData {
  theme: 'light' | 'dark' | 'system';
}

export interface CommandSetLocaleData {
  locale: string;
}

// ---------------------------------------------------------------
// Signing (uniform with the SDK `signing` prop)
// ---------------------------------------------------------------

export interface SignatureRequestData {
  fields: SignatureField[];
  mode: SignatureMode;
  banner?: string;
}

export interface SignatureRequestAckData {
  ok: boolean;
  code?: string;
}

export type SignatureFieldSignedData = SignedFieldPayload;
export type SignatureCompleteData = SignatureCompletePayload;

export interface SignatureCancelData {
  reason: CancelReason;
}

// ---------------------------------------------------------------
// Type guards
// ---------------------------------------------------------------

export function isCasualEnvelope(value: unknown): value is CasualEnvelope {
  if (!value || typeof value !== 'object') return false;
  const v = value as Record<string, unknown>;
  return (
    typeof v.type === 'string' &&
    v.type.startsWith('casual.') &&
    (v.app === 'docs' || v.app === 'sheet') &&
    v.v === 1 &&
    'data' in v
  );
}
