/**
 * Signing types — mirror the iframe-protocol envelopes from
 * `docs/internal/13-iframe-protocol.md` so the SDK and iframe
 * deliveries can hand the same payloads back and forth.
 *
 * Uniform across `app: 'docs' | 'sheet'`. Only the `anchor`
 * discriminator changes.
 */

export type SignatureMethod = 'drawn' | 'typed' | 'uploaded';

export type SignatureMode = 'sequential' | 'concurrent';

/**
 * Doc-anchored field — paragraph id from the editor's `w14:paraId`,
 * with an optional sub-paragraph search for placing the signature
 * inside a phrase rather than a whole block.
 */
export interface DocAnchor {
  kind: 'doc';
  paraId: string;
  search?: string;
}

/** Sheet-anchored field — `sheet` name + A1-style `cell` ref. */
export interface SheetAnchor {
  kind: 'sheet';
  sheet: string;
  cell: string;
}

export type SignatureAnchor = DocAnchor | SheetAnchor;

export interface SignatureField {
  /** Per-field id supplied by the host; echoed on every progress event. */
  fieldId: string;
  /** Label rendered next to the field — "Employee signature", etc. */
  label: string;
  /** Required fields must complete before `onComplete` fires. */
  required: boolean;
  /** Where the signature lands in the document. */
  anchor: SignatureAnchor;
  /** Allowed signature methods. */
  methods: SignatureMethod[];
  /** Optional signer identity the host knows about. */
  signer?: { name?: string; email?: string };
}

/** Payload emitted when a signer completes one field. */
export interface SignedFieldPayload {
  fieldId: string;
  method: SignatureMethod;
  /** Raw signature material — PNG for drawn, UTF-8 string-as-bytes for typed, host-attested bytes for uploaded. */
  bytes: ArrayBuffer;
  mime: string;
  /** Client wall-clock at completion; host typically pairs with a server timestamp. */
  signedAt: string;
  /** Optional placement hint a host generating a flat PDF can use. */
  placement?: { page: number; xPct: number; yPct: number };
}

/** Payload emitted when all required fields are signed. */
export interface SignatureCompletePayload {
  fieldIds: string[];
  /** Document bytes WITH stamps applied — note (v1): the editor
   *  returns the unmodified document and a `fields` map so the
   *  HOST stamps the final bytes. v2 lands editor-side stamping. */
  bytes: ArrayBuffer;
  fields: Record<string, SignedFieldPayload>;
}

export type CancelReason = 'signer_cancelled' | 'session_expired' | 'host_aborted';

/**
 * Configuration the host hands the editor when opening a signing
 * session. Maps 1:1 to the iframe `signature.request` envelope.
 */
export interface SigningSessionConfig {
  fields: SignatureField[];
  mode: SignatureMode;
  /** Optional banner the editor renders at the top of the pane. */
  banner?: string;
  /** Fires once per field as the signer completes it. */
  onFieldSigned?: (payload: SignedFieldPayload) => void | Promise<void>;
  /** Fires after every required field is signed. */
  onComplete?: (payload: SignatureCompletePayload) => void | Promise<void>;
  /** Fires when either side aborts the session. */
  onCancel?: (payload: { reason: CancelReason }) => void;
}
