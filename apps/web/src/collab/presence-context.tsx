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

import { createContext, useContext } from 'react';
import type { Identity, Peer } from './presence';

export type PresenceCtxValue = {
  /** Current local-user identity (name + derived color). Null until the
   *  driver has determined we're in a room and resolved the display name. */
  me: Identity | null;
  /** Active peers (excludes self). Sorted by clientId for stable render order. */
  peers: Peer[];
  /** Whether we should currently show the "set display name" prompt. */
  needsNamePrompt: boolean;
  /** Called by NamePrompt on submit. */
  setName: (name: string) => void;
  /** Called when the user dismisses the prompt — keep a generated name. */
  dismissNamePrompt: () => void;
};

export const PresenceContext = createContext<PresenceCtxValue>({
  me: null,
  peers: [],
  needsNamePrompt: false,
  setName: () => undefined,
  dismissNamePrompt: () => undefined,
});

export function usePresence(): PresenceCtxValue {
  return useContext(PresenceContext);
}
