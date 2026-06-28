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

import type ExcelJS from 'exceljs';
import type { IWorkbookData } from '@univerjs/core';

/**
 * xlsx-native `cell.note` ⇄ Univer `thread-comment` resource bridge.
 *
 * Comments in xlsx live on the cell itself (`xl/comments<N>.xml`).
 * Univer's thread-comment plugin stores them in a workbook-level
 * resource keyed by sheet id under the name below — same
 * registration pattern as defined-names. Each side of the round-trip
 * has to translate to the shape the other one expects:
 *
 *   xlsx cell.note (string | { texts: [{ text }] })
 *     ⇅
 *   { dataStream: "<text>\r\n" }  (Univer IDocumentBody minimal form)
 *
 * Why this exists separately from the parser/exporter modules:
 *   1. Two-sided code lives next to itself so the inverse stays
 *      obvious. The audit test fails the moment one side drifts.
 *   2. The `IThreadComment` row builder is non-trivial and reused
 *      verbatim on both sides — extracting it keeps parse-impl and
 *      export-impl from carrying near-duplicate shape definitions.
 */

// Mirrors vendor/univer/packages/thread-comment/src/controllers/tc-resource.controller.ts:25
// — the resource registration name the plugin's controller subscribes
// to at workbook load. Hard-coding instead of importing avoids pulling
// the plugin into the xlsx worker bundle (thread-comment ships UI deps
// that don't tree-shake cleanly in a Worker context).
export const THREAD_COMMENT_RESOURCE = 'SHEET_UNIVER_THREAD_COMMENT_PLUGIN';

// Minimal subset of Univer's `IThreadComment` we synthesise. The
// plugin populates the rest (children, mentions, resolved) on
// subsequent edits; xlsx-native comments don't carry any of them.
export type SynthComment = {
  id: string;
  threadId: string;
  ref: string;        // "B2" style
  dT: string;         // ISO date
  personId: string;
  text: { dataStream: string };
  unitId: string;
  subUnitId: string;
  // children: [] on import — Univer's controller tolerates this
  // because `addComment` of a root with no children is the normal
  // create path.
  children?: never[];
};

const PERSON_FROM_XLSX = 'imported';

function colToLetters(n: number): string {
  let out = '';
  let v = n;
  while (v >= 0) {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  }
  return out;
}

function refOf(row: number, column: number): string {
  return `${colToLetters(column)}${row + 1}`;
}

function noteToString(note: unknown): string | null {
  if (typeof note === 'string') return note.trim() ? note : null;
  if (note && typeof note === 'object') {
    // ExcelJS's expanded shape: { texts: [{ text }] }. Fall back to
    // `.text` if a single-string variant came back through some
    // odder code path.
    const texts = (note as { texts?: Array<{ text?: string }> }).texts;
    if (Array.isArray(texts)) {
      const joined = texts.map((t) => t?.text ?? '').join('');
      return joined.trim() ? joined : null;
    }
  }
  return null;
}

/**
 * Walk every worksheet and collect xlsx-native `cell.note` entries
 * into a Univer thread-comment resource payload. The synthesised IDs
 * and dates aren't stable — re-opening the same file produces new
 * ones — which is fine for the use cases we care about (comments
 * survive the round-trip; nobody is referring to them by id).
 */
export function readCommentsFromXlsx(
  wb: ExcelJS.Workbook,
  unitId: string,
  sheetIdForExcel: (excelId: number) => string,
): Record<string, SynthComment[]> {
  const out: Record<string, SynthComment[]> = {};
  let seq = 0;
  const nowIso = new Date().toISOString();

  for (const ws of wb.worksheets) {
    const subUnitId = sheetIdForExcel(ws.id);
    const bucket: SynthComment[] = [];
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        // ExcelJS exposes notes on cells with comments — the field is
        // missing entirely on the vast majority of cells, so iterating
        // every cell is cheap.
        const note = (cell as unknown as { note?: unknown }).note;
        const text = noteToString(note);
        if (!text) return;
        const id = `xc-${unitId}-${seq++}`;
        bucket.push({
          id,
          threadId: id,           // single-comment thread; thread id = root id
          ref: refOf(rowNumber - 1, colNumber - 1),
          dT: nowIso,
          personId: PERSON_FROM_XLSX,
          // Univer's body convention terminates with `\r\n` (one paragraph).
          // Without it the cursor in the panel jumps a column on first
          // open of the comment.
          text: { dataStream: `${text}\r\n` },
          unitId,
          subUnitId,
          children: [],
        });
      });
    });
    if (bucket.length > 0) out[subUnitId] = bucket;
  }

  return out;
}

/**
 * Merge a synthesised comment payload into the workbook's `resources`
 * array. Skipped when the workbook already carries a thread-comment
 * resource from our hidden sidecar (`__casual_sheets_resources__`) —
 * that one has the full plugin shape; we shouldn't clobber it with a
 * lossy xlsx-native re-derivation.
 */
export function mergeCommentsIntoResources(
  resources: IWorkbookData['resources'],
  comments: Record<string, SynthComment[]>,
): IWorkbookData['resources'] {
  if (Object.keys(comments).length === 0) return resources;
  const existing = resources?.find((r) => r.name === THREAD_COMMENT_RESOURCE);
  if (existing) return resources;
  const next = [...(resources ?? [])];
  next.push({
    name: THREAD_COMMENT_RESOURCE,
    data: JSON.stringify(comments),
  });
  return next;
}

/**
 * Read the thread-comment resource off a snapshot. Tolerant of older
 * / missing / malformed payloads — those cases return `{}` so the
 * exporter just skips writing notes rather than throwing on save.
 */
export function readCommentsFromSnapshot(
  data: IWorkbookData,
): Record<string, SynthComment[]> {
  const entry = data.resources?.find((r) => r.name === THREAD_COMMENT_RESOURCE);
  if (!entry?.data) return {};
  try {
    const parsed = JSON.parse(entry.data) as Record<string, unknown>;
    if (!parsed || typeof parsed !== 'object') return {};
    const out: Record<string, SynthComment[]> = {};
    for (const [sheetId, value] of Object.entries(parsed)) {
      if (!Array.isArray(value)) continue;
      const bucket: SynthComment[] = [];
      for (const c of value) {
        if (!c || typeof c !== 'object') continue;
        const obj = c as Record<string, unknown>;
        const text = (obj.text as { dataStream?: string } | undefined)?.dataStream;
        const ref = typeof obj.ref === 'string' ? obj.ref : null;
        if (!ref || typeof text !== 'string') continue;
        bucket.push(c as SynthComment);
      }
      if (bucket.length > 0) out[sheetId] = bucket;
    }
    return out;
  } catch {
    return {};
  }
}

/**
 * Extract the plain string from a Univer comment body. Strips the
 * trailing `\r\n` (or `\n`) that the body convention appends so the
 * xlsx-side note doesn't show as a blank line below the text in
 * Excel's pop-up.
 */
export function commentBodyToString(body: SynthComment['text']): string {
  const s = body?.dataStream ?? '';
  return s.replace(/[\r\n]+$/, '');
}

/**
 * Parse an "A1"-style ref back to a zero-based (row, col). Falls back
 * to (-1, -1) on malformed input — exporter callers should skip in
 * that case.
 */
export function refToRowCol(ref: string): { row: number; column: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) return { row: -1, column: -1 };
  const letters = m[1];
  let col = 0;
  for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
  return { row: Number(m[2]) - 1, column: col - 1 };
}
