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

import ExcelJS from 'exceljs';

/**
 * Regression guard for the shared-formula import-corruption fix in
 * `parse-impl.ts`.
 *
 * Excel autofill (dragging a formula down a column) emits `t="shared"` cells:
 * one master carries the formula, the slaves reference it by address. ExcelJS
 * surfaces a slave's `cell.value` as `{ sharedFormula: <masterAddress>, result }`
 * — `sharedFormula` is the MASTER ADDRESS, not a formula. The parser must read
 * the translated `cell.formula` getter; using the master address corrupts every
 * slave into `=<master>` and recalc/save destroys the whole column.
 *
 * This pins the exact ExcelJS contract the fix depends on. (parse-impl.ts can't
 * be imported under `node --test`: it pulls in @univerjs/core, whose built ESM
 * has no node-detectable named exports, so we assert the upstream behaviour the
 * inline fix relies on rather than the parser module itself.)
 */
test('exceljs Cell.formula translates shared-formula slaves to their own position', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.getCell('A1').value = 1;
  ws.getCell('A2').value = 2;
  ws.getCell('A3').value = 3;
  // B1 master, B2:B3 shared slaves of `A*2`.
  ws.fillFormula('B1:B3', 'A1*2', [2, 4, 6]);

  const buffer = await wb.xlsx.writeBuffer();
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(buffer);
  const ws2 = wb2.getWorksheet('Sheet1')!;

  for (const [addr, expected] of [
    ['B2', 'A2*2'],
    ['B3', 'A3*2'],
  ] as const) {
    const cell = ws2.getCell(addr);
    // Stored value is the master address — the exact shape the bug mis-read.
    assert.equal((cell.value as { sharedFormula: string }).sharedFormula, 'B1');
    // The getter gives the correct translated formula the parser must use.
    assert.equal(cell.formula, expected);
    assert.notEqual(cell.formula, 'B1');
  }
});
