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
 * Pluggable @mention data source for cell comments — pure layer.
 *
 * Univer's comment editor (the in-cell thread-comment popup) gets @-mention
 * autocomplete from `docs-mention-ui`, which lists candidates via the core
 * `IMentionIOService`. The default `MentionIOLocalService` is hardwired to a
 * single candidate — the current user — which is useless: you can't @mention
 * the people you're editing with.
 *
 * We override it with `CasualMentionIOService` (see `mention-io.ts`), which
 * reads candidates from a host-settable provider. The provider is the
 * pluggable seam the brief calls for:
 *   - embedded in another platform → the host sets it to its directory/users,
 *   - standalone → the app sets it to the live presence peers (+ self).
 *
 * This module is the @univerjs-value-free core (only `import type`, which is
 * erased at runtime) so it's unit-testable under the `node:test` + tsx runner,
 * which can't load @univerjs *value* modules. The enum-coupled service class
 * lives in `mention-io.ts`.
 */
import type { IListMentionParam, IListMentionResponse, MentionType } from '@univerjs/core';

/** One mentionable entity. `id` must be stable enough to resolve later. */
export type MentionCandidate = { id: string; label: string; icon?: string };

/** Resolves mention candidates for the current `@`-search term. */
export type MentionProvider = (search: string) => MentionCandidate[] | Promise<MentionCandidate[]>;

let provider: MentionProvider | null = null;

/**
 * Install (or clear, with `null`) the mention candidate provider. Last writer
 * wins — the host or the app sets this once. Safe to call before or after the
 * editor boots; the service reads it lazily on each `@`-trigger.
 */
export function setMentionProvider(fn: MentionProvider | null): void {
  provider = fn;
}

/** Current provider — exposed for tests / introspection. */
export function getMentionProvider(): MentionProvider | null {
  return provider;
}

/**
 * Case-insensitive substring filter applied to a candidate list. Strips the
 * leading `@` trigger — `docs-mention-ui` passes the search slice starting at
 * the `@` anchor (e.g. `"@gr"`), so the raw term includes it.
 */
export function filterMentionCandidates(
  candidates: MentionCandidate[],
  search: string,
): MentionCandidate[] {
  const q = search.replace(/^@+/, '').trim().toLowerCase();
  if (!q) return candidates;
  return candidates.filter((c) => c.label.toLowerCase().includes(q));
}

/**
 * Resolve the active provider into the `IMentionIOService.list` response shape.
 * `personType` is the numeric `MentionType.PERSON` value, passed in by the
 * service class so this module needs no @univerjs value import. A failing
 * provider degrades to an empty list — it must never break the editor.
 */
export async function resolveMentionList(
  params: IListMentionParam,
  personType: MentionType,
): Promise<IListMentionResponse> {
  const search = params.search ?? '';
  let candidates: MentionCandidate[] = [];
  if (provider) {
    try {
      candidates = await provider(search);
    } catch {
      candidates = [];
    }
  }
  const mentions = candidates.map((c) => ({
    objectType: personType,
    objectId: c.id,
    label: c.label,
    metadata: c.icon ? { icon: c.icon } : undefined,
  }));
  return {
    list: [{ type: personType, mentions, metadata: {}, title: 'PEOPLE' }],
    page: params.page,
    size: params.size,
    total: mentions.length,
  };
}
