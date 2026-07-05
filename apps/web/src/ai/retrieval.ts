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
 * BM25 lexical retrieval — pure TypeScript, no embeddings/model/download.
 * Vendored from the docs @casualoffice/docops retrieval package (the two repos
 * don't share a package yet). Runs JS-side so the local model never receives
 * the whole workbook. Keep in sync with docx-editor's packages/docops/src/retrieval.
 */

export interface RetrievalChunk {
  id: string;
  text: string;
  meta?: Record<string, unknown>;
}
export interface RetrievedChunk extends RetrievalChunk {
  score: number;
}
export interface RetrieveOptions {
  k?: number;
  charBudget?: number;
}
export interface RetrieveResult {
  chunks: RetrievedChunk[];
  truncated: boolean;
}

const K1 = 1.5;
const B = 0.75;
const DEFAULT_K = 5;
const DEFAULT_CHAR_BUDGET = 10_000;

const STOPWORDS = new Set(
  (
    'a an and are as at be but by for from has have i in is it its of on or that the to was were will with ' +
    'this these those what which who whom whose when where why how do does did can could should would may might'
  ).split(' '),
);

export function tokenize(text: string): string[] {
  return text
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter((t) => t.length > 1 && !STOPWORDS.has(t));
}

export function retrieve(
  chunks: RetrievalChunk[],
  query: string,
  options: RetrieveOptions = {},
): RetrieveResult {
  const k = options.k ?? DEFAULT_K;
  const charBudget = options.charBudget ?? DEFAULT_CHAR_BUDGET;

  const docs = chunks.map((chunk) => {
    const tokens = tokenize(chunk.text);
    const tf = new Map<string, number>();
    for (const tok of tokens) tf.set(tok, (tf.get(tok) ?? 0) + 1);
    return { chunk, tf, len: tokens.length };
  });
  const df = new Map<string, number>();
  for (const d of docs) for (const tok of d.tf.keys()) df.set(tok, (df.get(tok) ?? 0) + 1);
  const avgLen = docs.length ? docs.reduce((s, d) => s + d.len, 0) / docs.length : 0;
  const N = docs.length;

  const qTokens = [...new Set(tokenize(query))];
  const ranked = docs
    .map(({ chunk, tf, len }) => {
      let s = 0;
      for (const q of qTokens) {
        const f = tf.get(q);
        if (!f) continue;
        const n = df.get(q) ?? 0;
        const idf = Math.max(0, Math.log((N - n + 0.5) / (n + 0.5) + 1));
        const denom = f + K1 * (1 - B + (B * len) / (avgLen || 1));
        s += idf * ((f * (K1 + 1)) / denom);
      }
      return { ...chunk, score: s };
    })
    .filter((c) => c.score > 0)
    .sort((a, b) => b.score - a.score);

  const picked: RetrievedChunk[] = [];
  let used = 0;
  let truncated = false;
  for (const chunk of ranked) {
    if (picked.length >= k) {
      truncated = true;
      break;
    }
    if (used + chunk.text.length > charBudget) {
      truncated = true;
      continue;
    }
    picked.push(chunk);
    used += chunk.text.length;
  }
  return { chunks: picked, truncated };
}
