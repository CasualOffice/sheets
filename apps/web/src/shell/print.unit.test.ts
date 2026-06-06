/**
 * Print path guards — covers the cell-count cap added for #50.
 *
 * The renderer itself touches `document` + iframe APIs so a full
 * round-trip lives in the dev-server e2e (see
 * `tests/e2e/print-large-cap.spec.ts`). This unit test pins the
 * pure-input guard so future limit changes are deliberate.
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import { parsePrintArea, PRINT_CELL_LIMIT, printActiveSheet, type PrintOptions } from './print';

const DEFAULTS: PrintOptions = {
  orientation: 'portrait',
  margins: 'normal',
  printArea: null,
};

function makeApi(rows: number, cols: number) {
  return {
    getActiveWorkbook() {
      return {
        getName() {
          return 'Test';
        },
        getActiveSheet() {
          return {
            getSheet() {
              return {
                getLastRowWithContent: () => rows - 1,
                getLastColumnWithContent: () => cols - 1,
                getCell: () => ({ v: 'x' }),
                getColumnWidth: () => 88,
              };
            },
            getSheetName: () => 'Sheet1',
          };
        },
      };
    },
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;
}

test('refuses to print past PRINT_CELL_LIMIT', () => {
  // 500 × 250 = 125_000 cells — above the 100k cap.
  const result = printActiveSheet(makeApi(500, 250), DEFAULTS);
  assert.equal(result.ok, false);
  if (!result.ok) {
    assert.equal(result.reason, 'too-large');
    if (result.reason === 'too-large') {
      assert.equal(result.cellCount, 125_000);
      assert.equal(result.limit, PRINT_CELL_LIMIT);
    }
  }
});

test('reports empty when the sheet has no used range', () => {
  const result = printActiveSheet(makeApi(0, 0), DEFAULTS);
  // 0×0 = 1 cell (lastRow=lastCol=-1+1=0 vs startRow/startCol=0,
  // so the row/col counts are 0). lastRow < 0 short-circuits to
  // 'empty'.
  assert.equal(result.ok, false);
  if (!result.ok) assert.equal(result.reason, 'empty');
});

test('uses Print Area to shrink under the cap', () => {
  // Used range is 500×250 (too large), but Print Area limits it
  // down to A1:J10 = 100 cells. The size guard should pass; the
  // function then proceeds to DOM APIs and throws under node (no
  // `document`). What we're pinning here is that the guard does
  // not refuse for too-large.
  let thrown: Error | null = null;
  try {
    printActiveSheet(makeApi(500, 250), { ...DEFAULTS, printArea: 'A1:J10' });
  } catch (err) {
    thrown = err as Error;
  }
  assert.ok(
    thrown !== null && /document is not defined|createElement/.test(thrown.message),
    'expected the function to clear the size guard and reach DOM APIs',
  );
});

test('parsePrintArea handles A1:D20 and single A1', () => {
  assert.deepEqual(parsePrintArea('A1:D20'), {
    startRow: 0,
    startCol: 0,
    endRow: 19,
    endCol: 3,
  });
  assert.deepEqual(parsePrintArea('A1'), {
    startRow: 0,
    startCol: 0,
    endRow: 0,
    endCol: 0,
  });
  assert.equal(parsePrintArea('not-a-range'), null);
  assert.equal(parsePrintArea(''), null);
});
