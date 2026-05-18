import type { FUniver } from '@univerjs/core/facade';
import type { ChartType, ChartModel } from './types';
import { newChartId } from './types';

/**
 * Build a ChartModel from the active selection. The selection is the
 * source range; the chart anchors to a 10-row × 8-col block placed
 * just below the source so the chart doesn't cover its own data.
 *
 * Excel's Insert > Chart picks defaults the same way: source = the
 * current selection (or auto-detected contiguous range if a single
 * cell is active), output = roughly the right size, dropped below
 * the source. P1 will swap this for a real dialog with a range
 * picker + type picker; for P0 the menu item Just Works.
 *
 * Returns null when there's no active workbook/sheet or the
 * selection has fewer than 2 rows × 2 cols (no header + data).
 */
export function buildChartModelFromActiveSelection(
  api: FUniver,
  type: ChartType = 'bar',
): ChartModel | null {
  const wb = api.getActiveWorkbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ws = wb?.getActiveSheet() as any;
  const range = ws?.getActiveRange();
  if (!wb || !ws || !range) return null;

  const r = range.getRange();
  const rows = r.endRow - r.startRow + 1;
  const cols = r.endColumn - r.startColumn + 1;
  if (rows < 2 || cols < 2) return null;

  // Anchor: 10 rows × 8 cols, dropped one row below the source.
  const chartTop = r.endRow + 2;
  const chartLeft = r.startColumn;
  return {
    id: newChartId(),
    sheetId: ws.getSheetId(),
    source: {
      startRow: r.startRow,
      endRow: r.endRow,
      startColumn: r.startColumn,
      endColumn: r.endColumn,
    },
    pos: {
      startRow: chartTop,
      endRow: chartTop + 9,
      startColumn: chartLeft,
      endColumn: chartLeft + 7,
    },
    type,
  };
}
