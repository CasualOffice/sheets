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
 * SigningProvider — wraps a signing session in React context so
 * descendant components (SigningPane, capture surfaces, and the
 * editor's field-highlight decorations) all see the same snapshot.
 *
 * The pure state machine lives in `./controller.ts`; this file
 * just bridges it to React.
 */

import { createContext, useContext, useEffect, useMemo, useState, type ReactNode } from 'react';

import {
  createSigningController,
  type SigningController,
  type SigningSnapshot,
} from './controller';
import type {
  CancelReason,
  SignatureCompletePayload,
  SignedFieldPayload,
  SigningSessionConfig,
} from './types';

interface SigningContextValue {
  controller: SigningController;
  snapshot: SigningSnapshot;
  /** Convenience wrapper around controller.signField that also fires the host's onFieldSigned. */
  signField: (payload: SignedFieldPayload) => Promise<void>;
  /** Convenience wrapper around controller.complete + host's onComplete. */
  completeIfReady: () => Promise<void>;
  /** Convenience wrapper around controller.cancel + host's onCancel. */
  cancel: (reason: CancelReason) => void;
  /** Source-of-truth document bytes the editor renders. Captured at
   *  signing-session open; persists for the lifetime of the session. */
  baseDocumentBytes: ArrayBuffer | null;
}

const SigningContext = createContext<SigningContextValue | null>(null);

export interface SigningProviderProps {
  /** Active signing session config. When null, signing is off and
   *  children render unchanged. */
  session: SigningSessionConfig | null;
  /** Current document bytes the editor is rendering. Captured into
   *  the context so the eventual `complete` payload carries the
   *  right base buffer. */
  documentBytes: ArrayBuffer | null;
  children: ReactNode;
}

export function SigningProvider({ session, documentBytes, children }: SigningProviderProps) {
  if (!session) {
    return <>{children}</>;
  }
  return (
    <SigningProviderInner session={session} documentBytes={documentBytes}>
      {children}
    </SigningProviderInner>
  );
}

function SigningProviderInner({
  session,
  documentBytes,
  children,
}: {
  session: SigningSessionConfig;
  documentBytes: ArrayBuffer | null;
  children: ReactNode;
}) {
  // Controller is constructed once per session-config identity.
  // Hosts that swap sessions mid-tree must change the React `key`
  // on the provider — otherwise stale state would leak.
  const controller = useMemo(
    () => createSigningController(session.fields, session.mode),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [session.fields, session.mode],
  );

  const [snapshot, setSnapshot] = useState<SigningSnapshot>(() => controller.snapshot());

  useEffect(() => {
    const unsub = controller.subscribe(setSnapshot);
    return unsub;
  }, [controller]);

  const value = useMemo<SigningContextValue>(
    () => ({
      controller,
      snapshot,
      signField: async (payload) => {
        controller.signField(payload);
        await session.onFieldSigned?.(payload);
      },
      completeIfReady: async () => {
        if (!controller.snapshot().canComplete) return;
        const final = controller.complete();
        const completePayload: SignatureCompletePayload = {
          fieldIds: final.fields
            .map((f) => f.fieldId)
            .filter((id) => final.signed[id] !== undefined),
          bytes: documentBytes ?? new ArrayBuffer(0),
          fields: final.signed,
        };
        await session.onComplete?.(completePayload);
      },
      cancel: (reason) => {
        controller.cancel();
        session.onCancel?.({ reason });
      },
      baseDocumentBytes: documentBytes,
    }),
    [controller, snapshot, session, documentBytes],
  );

  return <SigningContext.Provider value={value}>{children}</SigningContext.Provider>;
}

/**
 * Hook for descendants of <SigningProvider>. Returns null when
 * no signing session is active — caller renders its non-signing
 * shape.
 */
export function useSigning(): SigningContextValue | null {
  return useContext(SigningContext);
}
