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
 * Signing state machine — pure with respect to React so bun-test
 * can exercise it without a renderer. The React wrappers
 * (SigningProvider / SigningPane) sit on top of this.
 *
 * Single source of truth for "which field is the signer on,
 * which fields are done, can we complete yet" — the React layer
 * never recomputes these from scratch.
 */

import type { SignedFieldPayload, SignatureField, SignatureMode } from './types';

export interface SigningSnapshot {
  fields: SignatureField[];
  mode: SignatureMode;
  /** Fields completed in order they were signed. */
  signed: Record<string, SignedFieldPayload>;
  /** Index into `fields` of the field the signer should focus on
   *  next. -1 once everything required is done. */
  activeFieldIndex: number;
  /** True once every required field has a payload. */
  canComplete: boolean;
  /** True once the controller has received a completion command. */
  isComplete: boolean;
  /** True if a cancel has been emitted. */
  isCancelled: boolean;
}

export interface SigningController {
  snapshot(): SigningSnapshot;
  subscribe(listener: (s: SigningSnapshot) => void): () => void;
  /** Record a signed field. Throws if the field id is unknown. */
  signField(payload: SignedFieldPayload): void;
  /** Move focus to a specific field (concurrent mode let users
   *  pick); in sequential mode the controller silently no-ops if
   *  the requested field isn't the next required. */
  focusField(fieldId: string): void;
  /** Mark the session complete. Returns the snapshot at the
   *  moment of completion; throws if not yet canComplete. */
  complete(): SigningSnapshot;
  /** Mark the session cancelled. Idempotent. */
  cancel(): void;
}

export function createSigningController(
  fields: SignatureField[],
  mode: SignatureMode,
): SigningController {
  if (fields.length === 0) {
    throw new Error('createSigningController: at least one field required');
  }
  const fieldIds = new Set(fields.map((f) => f.fieldId));
  if (fieldIds.size !== fields.length) {
    throw new Error('createSigningController: duplicate fieldId');
  }

  const signed: Record<string, SignedFieldPayload> = {};
  let activeFieldIndex = nextRequiredIndex(fields, signed);
  let isComplete = false;
  let isCancelled = false;
  const listeners = new Set<(s: SigningSnapshot) => void>();

  function emit() {
    const snap = snapshotInternal();
    for (const l of listeners) l(snap);
  }

  function snapshotInternal(): SigningSnapshot {
    return {
      fields,
      mode,
      signed: { ...signed },
      activeFieldIndex,
      canComplete: allRequiredSigned(fields, signed),
      isComplete,
      isCancelled,
    };
  }

  return {
    snapshot: snapshotInternal,
    subscribe(listener) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    signField(payload) {
      if (isComplete || isCancelled) return;
      if (!fieldIds.has(payload.fieldId)) {
        throw new Error(`signField: unknown fieldId ${payload.fieldId}`);
      }
      signed[payload.fieldId] = payload;
      activeFieldIndex = nextRequiredIndex(fields, signed);
      emit();
    },
    focusField(fieldId) {
      if (isComplete || isCancelled) return;
      const idx = fields.findIndex((f) => f.fieldId === fieldId);
      if (idx < 0) return;
      if (mode === 'sequential') {
        // Sequential mode: only the next-required field is focusable.
        const next = nextRequiredIndex(fields, signed);
        if (idx !== next) return;
      }
      activeFieldIndex = idx;
      emit();
    },
    complete() {
      if (!allRequiredSigned(fields, signed)) {
        throw new Error('complete: required fields are still unsigned');
      }
      isComplete = true;
      activeFieldIndex = -1;
      emit();
      return snapshotInternal();
    },
    cancel() {
      if (isComplete || isCancelled) return;
      isCancelled = true;
      activeFieldIndex = -1;
      emit();
    },
  };
}

// ---------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------

function allRequiredSigned(
  fields: SignatureField[],
  signed: Record<string, SignedFieldPayload>,
): boolean {
  return fields.every((f) => !f.required || signed[f.fieldId] !== undefined);
}

function nextRequiredIndex(
  fields: SignatureField[],
  signed: Record<string, SignedFieldPayload>,
): number {
  for (let i = 0; i < fields.length; i++) {
    if (fields[i].required && !signed[fields[i].fieldId]) return i;
  }
  // Either all required are done OR no required exists; pick the
  // first unsigned optional.
  for (let i = 0; i < fields.length; i++) {
    if (!signed[fields[i].fieldId]) return i;
  }
  return -1;
}
