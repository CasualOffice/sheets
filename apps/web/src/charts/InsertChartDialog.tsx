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

import { useMemo, useRef, useState } from 'react';
import { Dialog } from '../shell/Dialog';
import type { FUniver } from '@univerjs/core/facade';
import {
  CHART_FAMILY_OF,
  CHART_TYPE_LABEL,
  type ChartFamily,
  type ChartType,
} from './types';

type Props = {
  api: FUniver;
  /** A1 reference to seed the source-range input. Pulled from the active
   *  selection by the caller so the dialog opens already filled in. */
  defaultSourceA1: string;
  /** Pre-selected chart type. Defaults to `'column'` (Excel's
   *  Insert > Chart default — Clustered Column). Set to the chart's
   *  current type when used for "Change chart type". */
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

type FamilyDef = {
  id: ChartFamily;
  label: string;
  icon: string;
  subtypes: { id: ChartType; label: string }[];
};

/**
 * Excel chart catalog, family + subtype hierarchy. Order mirrors
 * Excel's Insert > Chart panel so muscle memory transfers.
 */
const FAMILIES: FamilyDef[] = [
  {
    id: 'column',
    label: 'Column',
    icon: 'bar_chart',
    subtypes: [
      { id: 'column', label: CHART_TYPE_LABEL.column },
      { id: 'column-stacked', label: CHART_TYPE_LABEL['column-stacked'] },
      { id: 'column-stacked-100', label: CHART_TYPE_LABEL['column-stacked-100'] },
    ],
  },
  {
    id: 'bar',
    label: 'Bar',
    icon: 'align_horizontal_left',
    subtypes: [
      { id: 'bar', label: CHART_TYPE_LABEL.bar },
      { id: 'bar-stacked', label: CHART_TYPE_LABEL['bar-stacked'] },
      { id: 'bar-stacked-100', label: CHART_TYPE_LABEL['bar-stacked-100'] },
    ],
  },
  {
    id: 'line',
    label: 'Line',
    icon: 'show_chart',
    subtypes: [
      { id: 'line', label: CHART_TYPE_LABEL.line },
      { id: 'line-stacked', label: CHART_TYPE_LABEL['line-stacked'] },
    ],
  },
  {
    id: 'area',
    label: 'Area',
    icon: 'area_chart',
    subtypes: [
      { id: 'area', label: CHART_TYPE_LABEL.area },
      { id: 'area-stacked', label: CHART_TYPE_LABEL['area-stacked'] },
    ],
  },
  {
    id: 'pie',
    label: 'Pie',
    icon: 'pie_chart',
    subtypes: [
      { id: 'pie', label: CHART_TYPE_LABEL.pie },
      { id: 'doughnut', label: CHART_TYPE_LABEL.doughnut },
    ],
  },
  {
    id: 'scatter',
    label: 'Scatter',
    icon: 'scatter_plot',
    subtypes: [{ id: 'scatter', label: CHART_TYPE_LABEL.scatter }],
  },
];

/**
 * Excel-style "Insert chart" dialog. Two-panel layout:
 *
 *   - Left: family list (Column / Bar / Line / Area / Pie / Scatter).
 *     Clicking switches the visible subtypes.
 *   - Right top: subtype thumbnails with hover-to-preview labels.
 *   - Right bottom: source-range input, pre-filled from the active
 *     selection so the typical flow is two clicks.
 *
 * Range parsing leans on `sheet.getRange(a1)` — same A1 parser as the
 * Name Box — so anything the formula bar accepts works here too
 * (B5, A1:C4, B:B). Below 2 rows × 2 cols we refuse with an inline
 * hint instead of producing an empty chart.
 */
export function InsertChartDialog({
  api,
  defaultSourceA1,
  initialType = 'column',
  title = 'Insert chart',
  confirmLabel = 'Insert',
  onCancel,
  onConfirm,
}: Props) {
  const [type, setType] = useState<ChartType>(initialType);
  const [family, setFamily] = useState<ChartFamily>(CHART_FAMILY_OF[initialType]);
  const [sourceA1, setSourceA1] = useState(defaultSourceA1);
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const activeSheet = useMemo(() => {
    const wb = api.getActiveWorkbook();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return wb?.getActiveSheet() as any;
  }, [api]);

  const currentFamily = FAMILIES.find((f) => f.id === family) ?? FAMILIES[0];

  const pickFamily = (id: ChartFamily) => {
    setFamily(id);
    const fam = FAMILIES.find((f) => f.id === id);
    if (fam && !fam.subtypes.some((s) => s.id === type)) {
      // Switched family — pick its first subtype.
      setType(fam.subtypes[0].id);
    }
  };

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
        <div className="insert-chart__catalog">
          <ul
            className="insert-chart__families"
            role="tablist"
            aria-label="Chart family"
            data-testid="insert-chart-families"
          >
            {FAMILIES.map((f) => (
              <li key={f.id}>
                <button
                  type="button"
                  role="tab"
                  aria-selected={family === f.id}
                  className={`insert-chart__family${family === f.id ? ' insert-chart__family--active' : ''}`}
                  data-testid={`insert-chart-family-${f.id}`}
                  onClick={() => pickFamily(f.id)}
                >
                  <span className="material-symbols-outlined" aria-hidden="true">
                    {f.icon}
                  </span>
                  <span>{f.label}</span>
                </button>
              </li>
            ))}
          </ul>
          <div className="insert-chart__subtypes" role="radiogroup" aria-label="Chart subtype">
            {currentFamily.subtypes.map((s) => (
              <label
                key={s.id}
                className={`insert-chart__subtype${type === s.id ? ' insert-chart__subtype--active' : ''}`}
                data-testid={`insert-chart-type-${s.id}`}
                title={s.label}
              >
                <input
                  type="radio"
                  name="chart-subtype"
                  value={s.id}
                  checked={type === s.id}
                  onChange={() => setType(s.id)}
                />
                <SubtypeThumbnail type={s.id} />
                <span className="insert-chart__subtype-label">{s.label}</span>
              </label>
            ))}
          </div>
        </div>

        <fieldset className="insert-chart__group">
          <legend className="insert-chart__legend">Source data</legend>
          {/* Error elevates ABOVE the input so the user's eye lands
              on it before the field they're correcting. The previous
              "small red text below the input" layout was easy to
              miss (audit finding 1.2); a proper banner with icon +
              role=alert + aria-live=assertive makes it impossible
              to skip. Screen readers announce on appearance. */}
          {error && (
            <div
              className="insert-chart__error"
              data-testid="insert-chart-error"
              role="alert"
              aria-live="assertive"
            >
              <span className="insert-chart__error-icon" aria-hidden="true">!</span>
              <span>{error}</span>
            </div>
          )}
          <input
            ref={inputRef}
            type="text"
            className={`insert-chart__range${error ? ' insert-chart__range--error' : ''}`}
            data-testid="insert-chart-range"
            value={sourceA1}
            spellCheck={false}
            aria-invalid={Boolean(error)}
            aria-describedby={error ? 'insert-chart-error-desc' : undefined}
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
          <p className="insert-chart__hint" id="insert-chart-error-desc">
            First row is used as series labels and the first column as category
            labels — same convention as Excel's Insert &gt; Chart.
          </p>
        </fieldset>
      </div>
    </Dialog>
  );
}

/**
 * Inline SVG thumbnail for each subtype. ECharts canvases would be
 * too heavy to mount one per option; small static SVGs read at a
 * glance and match the Excel preview tiles' role.
 */
function SubtypeThumbnail({ type }: { type: ChartType }) {
  return (
    <svg
      className="insert-chart__thumb"
      viewBox="0 0 60 40"
      width={60}
      height={40}
      aria-hidden="true"
    >
      {renderThumbBody(type)}
    </svg>
  );
}

function renderThumbBody(type: ChartType): JSX.Element {
  const stroke = 'currentColor';
  switch (type) {
    case 'column':
      return (
        <g fill={stroke}>
          <rect x="6" y="20" width="6" height="14" />
          <rect x="16" y="12" width="6" height="22" />
          <rect x="26" y="8"  width="6" height="26" />
          <rect x="36" y="16" width="6" height="18" />
          <rect x="46" y="22" width="6" height="12" />
        </g>
      );
    case 'column-stacked':
      return (
        <g>
          <rect x="10" y="22" width="10" height="12" fill={stroke} />
          <rect x="10" y="12" width="10" height="10" fill={stroke} opacity="0.6" />
          <rect x="26" y="16" width="10" height="18" fill={stroke} />
          <rect x="26" y="8"  width="10" height="8"  fill={stroke} opacity="0.6" />
          <rect x="42" y="20" width="10" height="14" fill={stroke} />
          <rect x="42" y="10" width="10" height="10" fill={stroke} opacity="0.6" />
        </g>
      );
    case 'column-stacked-100':
      return (
        <g>
          <rect x="10" y="20" width="10" height="14" fill={stroke} />
          <rect x="10" y="6"  width="10" height="14" fill={stroke} opacity="0.6" />
          <rect x="26" y="14" width="10" height="20" fill={stroke} />
          <rect x="26" y="6"  width="10" height="8"  fill={stroke} opacity="0.6" />
          <rect x="42" y="24" width="10" height="10" fill={stroke} />
          <rect x="42" y="6"  width="10" height="18" fill={stroke} opacity="0.6" />
        </g>
      );
    case 'bar':
      return (
        <g fill={stroke}>
          <rect x="6" y="6"  width="22" height="5" />
          <rect x="6" y="14" width="36" height="5" />
          <rect x="6" y="22" width="48" height="5" />
          <rect x="6" y="30" width="28" height="5" />
        </g>
      );
    case 'bar-stacked':
      return (
        <g>
          <rect x="6"  y="10" width="20" height="6" fill={stroke} />
          <rect x="26" y="10" width="16" height="6" fill={stroke} opacity="0.6" />
          <rect x="6"  y="20" width="14" height="6" fill={stroke} />
          <rect x="20" y="20" width="28" height="6" fill={stroke} opacity="0.6" />
          <rect x="6"  y="30" width="30" height="6" fill={stroke} />
          <rect x="36" y="30" width="14" height="6" fill={stroke} opacity="0.6" />
        </g>
      );
    case 'bar-stacked-100':
      return (
        <g>
          <rect x="6"  y="10" width="28" height="6" fill={stroke} />
          <rect x="34" y="10" width="20" height="6" fill={stroke} opacity="0.6" />
          <rect x="6"  y="20" width="14" height="6" fill={stroke} />
          <rect x="20" y="20" width="34" height="6" fill={stroke} opacity="0.6" />
          <rect x="6"  y="30" width="40" height="6" fill={stroke} />
          <rect x="46" y="30" width="8"  height="6" fill={stroke} opacity="0.6" />
        </g>
      );
    case 'line':
      return (
        <g fill="none" stroke={stroke} strokeWidth="2">
          <polyline points="6,30 18,18 30,24 42,10 54,16" />
          <circle cx="6"  cy="30" r="2" fill={stroke} />
          <circle cx="18" cy="18" r="2" fill={stroke} />
          <circle cx="30" cy="24" r="2" fill={stroke} />
          <circle cx="42" cy="10" r="2" fill={stroke} />
          <circle cx="54" cy="16" r="2" fill={stroke} />
        </g>
      );
    case 'line-stacked':
      return (
        <g fill="none" strokeWidth="2">
          <polyline points="6,32 18,28 30,30 42,24 54,26" stroke={stroke} />
          <polyline points="6,20 18,12 30,18 42,8 54,14" stroke={stroke} opacity="0.6" />
        </g>
      );
    case 'area':
      return (
        <g>
          <polygon
            points="6,34 6,24 18,14 30,18 42,8 54,12 54,34"
            fill={stroke}
            opacity="0.4"
          />
          <polyline
            points="6,24 18,14 30,18 42,8 54,12"
            fill="none"
            stroke={stroke}
            strokeWidth="2"
          />
        </g>
      );
    case 'area-stacked':
      return (
        <g>
          <polygon
            points="6,34 6,24 18,18 30,22 42,16 54,20 54,34"
            fill={stroke}
            opacity="0.6"
          />
          <polygon
            points="6,24 18,18 30,22 42,16 54,20 54,12 42,6 30,12 18,8 6,14"
            fill={stroke}
            opacity="0.3"
          />
        </g>
      );
    case 'pie':
      return (
        <g>
          <circle cx="30" cy="20" r="14" fill={stroke} opacity="0.4" />
          <path d="M30 20 L30 6 A14 14 0 0 1 42 24 Z" fill={stroke} />
        </g>
      );
    case 'doughnut':
      return (
        <g>
          <circle cx="30" cy="20" r="14" fill={stroke} opacity="0.4" />
          <path d="M30 20 L30 6 A14 14 0 0 1 42 24 Z" fill={stroke} />
          <circle cx="30" cy="20" r="7" fill="white" />
        </g>
      );
    case 'scatter':
      return (
        <g fill={stroke}>
          <circle cx="10" cy="30" r="2" />
          <circle cx="18" cy="22" r="2" />
          <circle cx="24" cy="26" r="2" />
          <circle cx="30" cy="14" r="2" />
          <circle cx="36" cy="20" r="2" />
          <circle cx="42" cy="10" r="2" />
          <circle cx="48" cy="16" r="2" />
          <circle cx="14" cy="14" r="2" />
          <circle cx="50" cy="28" r="2" />
        </g>
      );
  }
}
