/**
 * Signing — document-signature pipeline.
 *
 * Surfaces:
 *   - SigningProvider + useSigning — React context wrapping the
 *     pure controller.
 *   - SigningPane — floating sidebar walking the signer through
 *     fields.
 *   - DrawnSignaturePad / TypedSignatureField / UploadedSignatureField
 *     — capture surfaces emitting { bytes, mime } pairs.
 *
 * Types mirror the iframe envelopes from
 * `docs/internal/13-iframe-protocol.md` — same shape across docs
 * and sheet.
 */

export { SigningProvider, useSigning, type SigningProviderProps } from './SigningProvider';
export { SigningPane, type SigningPaneProps } from './SigningPane';
export {
  DrawnSignaturePad,
  TypedSignatureField,
  UploadedSignatureField,
  type CapturedSignature,
  type DrawnSignaturePadProps,
  type TypedSignatureFieldProps,
  type UploadedSignatureFieldProps,
} from './captures';
export {
  createSigningController,
  type SigningController,
  type SigningSnapshot,
} from './controller';
export type {
  CancelReason,
  DocAnchor,
  SheetAnchor,
  SignatureAnchor,
  SignatureCompletePayload,
  SignatureField,
  SignatureMethod,
  SignatureMode,
  SignedFieldPayload,
  SigningSessionConfig,
} from './types';
