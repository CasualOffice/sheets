import { useMemo, useRef, useState } from 'react';
import { Dialog } from '../shell/Dialog';
import type { FUniver } from '@univerjs/core/facade';
import type { ChartType } from './types';

type Props = {
  api: FUniver;
  /** A1 reference to seed the source-range input. Pulled from the active
   *  selection by the caller so the dialog opens already filled in. */
  defaultSourceA1: string;
  /** Pre-selected chart type. Defaults to `'bar'` (Excel's Insert > Chart
   *  default). Set to the chart's current type when used for "Change
   *  chart type". */
  initialType?: ChartType;
  /** Dialog title — defaults to "Insert chart". */
  title?: string;
  /** Primary action label — defaults to "Insert". */
  confirmLabel?: string;
  onCancel: () => void;
  onConfirm: (args: {
    source: { startRow: number; endRow: number; startColumn: number; endColumn: number };
    type: ChartType;
  }) => void;
};

type TypeChoice = { id: ChartType; label: string; icon: string };

const TYPES: TypeChoice[] = [
  { id: 'bar', label: 'Column', icon: 'bar_chart' },
  { id: 'line', label: 'Line', icon: 'show_chart' },
  { id: 'pie', label: 'Pie', icon: 'pie_chart' },
  { id: 'scatter', label: 'Scatter', icon: 'scatter_plot' },
];

/**
 * Excel-style "Insert chart" dialog. Two inputs:
 *   - Chart type (segmented control, defaults to Column).
 *   - Source data range, as an A1 reference. Pre-filled from the active
 *     selection so the typical "select range → Insert → Chart" flow is
 *     two clicks.
 *
 * Parsing leans on `sheet.getRange(a1)` — same A1 parser as the Name Box
 * — so anything the formula bar accepts works here too (B5, A1:C4, B:B).
 * Below 2 rows × 2 cols the chart has no header + data axes, so the
 * dialog refuses with an inline hint instead of producing an empty chart.
 */
export function InsertChartDialog({
  api,
  defaultSourceA1,
  initialType = 'bar',
  title = 'Insert chart',
  confirmLabel = 'Insert',
  onCancel,
  onConfirm,
}: Props) {
  const [type, setType] = useState<ChartType>(initialType);
  const [sourceA1, setSourceA1] = useState(defaultSourceA1);
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const activeSheet = useMemo(() => {
    const wb = api.getActiveWorkbook();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return wb?.getActiveSheet() as any;
  }, [api]);

  const confirm = () => {
    const trimmed = sourceA1.trim();
    if (!trimmed) {
      setError('Pick a source range.');
      inputRef.current?.focus();
      return;
    }
    let range: { startRow: number; endRow: number; startColumn: number; endColumn: number } | null = null;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const r = activeSheet?.getRange(trimmed) as any;
      const raw = r?.getRange?.();
      if (raw) {
        range = {
          startRow: raw.startRow,
          endRow: raw.endRow,
          startColumn: raw.startColumn,
          endColumn: raw.endColumn,
        };
      }
    } catch {
      /* fall through */
    }
    if (!range) {
      setError("That doesn't look like a valid range — try A1:C4.");
      inputRef.current?.focus();
      return;
    }
    const rows = range.endRow - range.startRow + 1;
    const cols = range.endColumn - range.startColumn + 1;
    if (rows < 2 || cols < 2) {
      setError('Source needs at least a header row + a data row, and a label column + a value column.');
      inputRef.current?.focus();
      return;
    }
    onConfirm({ source: range, type });
  };

  return (
    <Dialog
      title={title}
      onClose={onCancel}
      data-testid="insert-chart-dialog"
      footer={
        <>
          <button
            type="button"
            className="btn-secondary"
            data-testid="insert-chart-cancel"
            onClick={onCancel}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="insert-chart-confirm"
            onClick={confirm}
          >
            {confirmLabel}
          </button>
        </>
      }
    >
      <div className="insert-chart">
        <fieldset className="insert-chart__group">
          <legend className="insert-chart__legend">Chart type</legend>
          <div
            className="insert-chart__types"
            role="radiogroup"
            aria-label="Chart type"
          >
            {TYPES.map((t) => (
              <label
                key={t.id}
                className={`insert-chart__type${type === t.id ? ' insert-chart__type--active' : ''}`}
                data-testid={`insert-chart-type-${t.id}`}
              >
                <input
                  type="radio"
                  name="chart-type"
                  value={t.id}
                  checked={type === t.id}
                  onChange={() => setType(t.id)}
                />
                <span className="material-symbols-outlined">{t.icon}</span>
                <span>{t.label}</span>
              </label>
            ))}
          </div>
        </fieldset>

        <fieldset className="insert-chart__group">
          <legend className="insert-chart__legend">Source data</legend>
          <input
            ref={inputRef}
            type="text"
            className="insert-chart__range"
            data-testid="insert-chart-range"
            value={sourceA1}
            spellCheck={false}
            onChange={(e) => {
              setSourceA1(e.target.value);
              if (error) setError(null);
            }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                confirm();
              }
            }}
            placeholder="A1:C4"
          />
          {error && (
            <div className="insert-chart__error" data-testid="insert-chart-error">
              {error}
            </div>
          )}
          <p className="insert-chart__hint">
            First row is used as series labels and the first column as category
            labels — same convention as Excel's Insert &gt; Chart.
          </p>
        </fieldset>
      </div>
    </Dialog>
  );
}
