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
 * Read @mentions out of a comment's rich-text body.
 *
 * A mention is encoded as a custom range with `rangeType === CustomRangeType.
 * MENTION`; the mentioned label is the `dataStream` slice it spans (see the
 * fork's `transformTextNodes2Document`). We surface "mentions you" in the
 * comments panel by matching those labels against the local display name.
 *
 * Pure + @univerjs-value-free (the MENTION enum value is passed in) so it's
 * unit-testable under the node:test + tsx runner. Matching is by name — the
 * shared identity across peers — since per-client mention ids aren't stable
 * (precise cross-user identity is Casual Drive's job).
 */

/** Minimal shape of the bits of `IDocumentBody` we read. */
export type MentionBody = {
  dataStream?: string;
  customRanges?: Array<{
    rangeType?: number;
    startIndex?: number;
    endIndex?: number;
  }>;
};

/**
 * Labels of every mention in the body, with the leading `@` trigger stripped.
 * `mentionType` is `CustomRangeType.MENTION` (passed in to keep this pure).
 */
export function extractMentionLabels(body: MentionBody | undefined, mentionType: number): string[] {
  const stream = body?.dataStream ?? '';
  const ranges = body?.customRanges ?? [];
  const labels: string[] = [];
  for (const r of ranges) {
    if (r.rangeType !== mentionType) continue;
    if (typeof r.startIndex !== 'number' || typeof r.endIndex !== 'number') continue;
    const raw = stream.slice(r.startIndex, r.endIndex + 1);
    const label = raw.replace(/^@+/, '').trim();
    if (label) labels.push(label);
  }
  return labels;
}

/** True if the body @mentions `name` (case-insensitive). Empty name → false. */
export function commentMentionsName(
  body: MentionBody | undefined,
  mentionType: number,
  name: string | null | undefined,
): boolean {
  const target = name?.trim().toLowerCase();
  if (!target) return false;
  return extractMentionLabels(body, mentionType).some((l) => l.toLowerCase() === target);
}
