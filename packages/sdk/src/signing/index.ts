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
