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

import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import {
  bodyToExcelRichText,
  excelRichTextToBody,
  type ExcelRichRun,
  type RichBody,
} from './rich-text.js';

test('excelRichTextToBody: builds textRuns + dataStream from styled runs', () => {
  const runs: ExcelRichRun[] = [
    { text: 'Hi', font: { bold: true } },
    { text: 'there', font: {} },
  ];
  const body = excelRichTextToBody(runs)!;
  assert.equal(body.dataStream, 'Hithere\r\n');
  assert.equal(body.textRuns!.length, 2);
  assert.deepEqual(body.textRuns![0], { st: 0, ed: 2, ts: { bl: 1 } });
  assert.deepEqual(body.textRuns![1], { st: 2, ed: 7 }); // no style → no ts
  assert.deepEqual(body.paragraphs, [{ startIndex: 7 }]);
  assert.deepEqual(body.sectionBreaks, [{ startIndex: 8 }]);
});

test('excelRichTextToBody: maps font name/size/italic/underline/color', () => {
  const body = excelRichTextToBody([
    {
      text: 'X',
      font: { name: 'Arial', size: 14, italic: true, underline: true, color: { argb: 'FFFF0000' } },
    },
  ])!;
  assert.deepEqual(body.textRuns![0].ts, {
    ff: 'Arial',
    fs: 14,
    it: 1,
    ul: { s: 1 },
    cl: { rgb: '#ff0000' },
  });
});

test('excelRichTextToBody: returns undefined when no run carries formatting', () => {
  assert.equal(excelRichTextToBody([{ text: 'a' }, { text: 'b' }]), undefined);
  assert.equal(excelRichTextToBody([]), undefined);
});

test('bodyToExcelRichText: emits runs with fonts from textRuns', () => {
  const body: RichBody = {
    dataStream: 'Hithere\r\n',
    textRuns: [
      { st: 0, ed: 2, ts: { bl: 1 } },
      { st: 2, ed: 7, ts: {} },
    ],
  };
  const runs = bodyToExcelRichText(body)!;
  assert.deepEqual(runs, [{ text: 'Hi', font: { bold: true } }, { text: 'there' }]);
});

test('bodyToExcelRichText: fills gaps between runs so no text is lost', () => {
  const body: RichBody = {
    dataStream: 'abcdef\r\n',
    textRuns: [{ st: 2, ed: 4, ts: { it: 1 } }], // only "cd" styled
  };
  const runs = bodyToExcelRichText(body)!;
  assert.deepEqual(runs, [{ text: 'ab' }, { text: 'cd', font: { italic: true } }, { text: 'ef' }]);
});

test('bodyToExcelRichText: undefined when there is no real per-run formatting', () => {
  assert.equal(
    bodyToExcelRichText({ dataStream: 'abc\r\n', textRuns: [{ st: 0, ed: 3 }] }),
    undefined,
  );
  assert.equal(bodyToExcelRichText({ dataStream: 'abc\r\n' }), undefined);
  assert.equal(bodyToExcelRichText(undefined), undefined);
});

test('round-trip: excel → body → excel preserves text + bold/italic', () => {
  const original: ExcelRichRun[] = [
    { text: 'Bold', font: { bold: true } },
    { text: ' and ' },
    { text: 'italic', font: { italic: true } },
  ];
  const body = excelRichTextToBody(original)!;
  const back = bodyToExcelRichText(body)!;
  assert.equal(back.map((r) => r.text).join(''), 'Bold and italic');
  assert.equal(back[0].font?.bold, true);
  assert.equal(back[2].font?.italic, true);
});
