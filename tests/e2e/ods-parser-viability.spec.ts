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

import { expect, test } from '@playwright/test';
import { createRequire } from 'node:module';
import * as path from 'node:path';

const require = createRequire(import.meta.url);
// eslint-disable-next-line @typescript-eslint/no-var-requires
const XLSX = require(require.resolve('@e965/xlsx', { paths: [path.join(process.cwd(), 'apps/web')] }));

test.describe('ods parser viability probes', () => {
  test('current SheetJS ODS read path does not surface row or column metadata', async () => {
    const ws = {
      A1: { t: 's', v: 'Header' },
      '!cols': [{ wch: 20, hidden: true }],
      '!rows': [{ hpt: 24, hidden: true }],
      '!ref': 'A1',
    };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');

    const out = XLSX.write(wb, { type: 'buffer', bookType: 'ods' });
    const back = XLSX.read(out, { type: 'buffer', cellNF: true, cellFormula: true });

    expect(back.Sheets.Data['!cols']).toBeUndefined();
    expect(back.Sheets.Data['!rows']).toBeUndefined();
  });

  test('current SheetJS ODS read path does not surface general font/fill/alignment style objects', async () => {
    const ws = {
      A1: {
        t: 's',
        v: 'Styled',
        s: {
          font: { bold: true, color: { rgb: 'FF0000' } },
          fill: { patternType: 'solid', fgColor: { rgb: '00FF00' } },
          alignment: { horizontal: 'center', vertical: 'center' },
        },
      },
      '!ref': 'A1',
    };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');

    const out = XLSX.write(wb, { type: 'buffer', bookType: 'ods', cellStyles: true });
    const back = XLSX.read(out, { type: 'buffer', cellStyles: true, cellNF: true, cellFormula: true });

    expect(back.Sheets.Data.A1?.s).toBeUndefined();
  });
});
