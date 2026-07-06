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
 * On-device workspace RAG (north-star O2) for Sheets — mirror of the docs
 * WorkspaceIndex + workspaceStore. BM25 retrieval across MANY documents (the
 * user's local folder); the host (desktop shell) extracts plain text from each
 * file and pushes it here, and search_workspace retrieves + cites across all of
 * them. One workspace shared across the app; nothing leaves the machine.
 */

import { retrieve, type RetrievalChunk } from './retrieval';

export interface WorkspaceDoc {
  id: string;
  name: string;
  text: string;
}
export interface WorkspaceHit {
  docId: string;
  docName: string;
  snippet: string;
  score: number;
}
export interface WorkspaceSearchResult {
  hits: WorkspaceHit[];
  sources: { docId: string; docName: string }[];
  truncated: boolean;
}

const CHUNK_CHARS = 1800;
const SNIPPET_CHARS = 700;

function chunkText(text: string): string[] {
  const paras = text
    .split(/\n\s*\n/)
    .map((p) => p.trim())
    .filter(Boolean);
  const chunks: string[] = [];
  let buf = '';
  for (const para of paras) {
    if (buf && buf.length + para.length > CHUNK_CHARS) {
      chunks.push(buf);
      buf = '';
    }
    buf = buf ? `${buf}\n${para}` : para;
    if (buf.length >= CHUNK_CHARS) {
      chunks.push(buf);
      buf = '';
    }
  }
  if (buf) chunks.push(buf);
  return chunks;
}

export class WorkspaceIndex {
  private chunks: RetrievalChunk[] = [];
  private readonly docs = new Map<string, string>();

  add(doc: WorkspaceDoc): void {
    this.remove(doc.id);
    this.docs.set(doc.id, doc.name);
    chunkText(doc.text).forEach((text, i) => {
      this.chunks.push({ id: `${doc.id}#${i}`, text, meta: { docId: doc.id, docName: doc.name } });
    });
  }

  remove(docId: string): void {
    if (!this.docs.has(docId)) return;
    this.docs.delete(docId);
    this.chunks = this.chunks.filter((c) => (c.meta as { docId: string }).docId !== docId);
  }

  get size(): number {
    return this.docs.size;
  }

  search(query: string, k = 6): WorkspaceSearchResult {
    const result = retrieve(this.chunks, query, { k });
    const hits: WorkspaceHit[] = result.chunks.map((c) => {
      const meta = c.meta as { docId: string; docName: string };
      return {
        docId: meta.docId,
        docName: meta.docName,
        snippet: c.text.slice(0, SNIPPET_CHARS),
        score: Math.round(c.score * 100) / 100,
      };
    });
    const seen = new Set<string>();
    const sources: { docId: string; docName: string }[] = [];
    for (const h of hits) {
      if (seen.has(h.docId)) continue;
      seen.add(h.docId);
      sources.push({ docId: h.docId, docName: h.docName });
    }
    return { hits, sources, truncated: result.truncated };
  }
}

// ── Shared process-level workspace (one folder, all editor instances) ────────

let shared: WorkspaceIndex | null = null;

export function setWorkspaceDocs(docs: WorkspaceDoc[]): void {
  if (!docs.length) {
    shared = null;
    return;
  }
  const idx = new WorkspaceIndex();
  for (const doc of docs) idx.add(doc);
  shared = idx;
}

export function clearWorkspace(): void {
  shared = null;
}

export function getSharedWorkspace(): WorkspaceIndex | null {
  return shared && shared.size > 0 ? shared : null;
}
