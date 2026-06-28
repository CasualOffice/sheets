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

import { useState } from 'react';
import { Dialog } from '../shell/Dialog';
import {
  CHART_FAMILY_OF,
  LEGEND_POSITIONS,
  PALETTES,
  PALETTE_LABELS,
  mergeFormat,
  type ChartFormat,
  type ChartModel,
  type ChartPalette,
} from './types';

type Props = {
  model: ChartModel;
  /** Optional list of series names — when present, the dialog renders
   *  a per-series colour-override picker below the palette section.
   *  Pulled by the caller from the chart's source range. */
  seriesNames?: string[];
  onCancel: () => void;
  onConfirm: (next: ChartFormat) => void;
};

/**
 * Excel's "Format Chart Area" pane, compressed into a single dialog
 * tailored to what we currently render. Covers:
 *
 *   - Chart title (text + show/hide).
 *   - Legend position (Bottom / Top / Left / Right / None).
 *   - X-axis title text + Y-axis title text (axis families only).
 *   - Major gridlines on/off (axis families only).
 *   - Data labels on/off.
 *   - Colour palette (Office / Mono / Vivid / Pastel).
 *
 * The dialog confirms with a merged `ChartFormat`; ChartContextMenu
 * applies it to `model.format` via `update()`.
 *
 * The chart title text is `model.title` (auto-named "Chart N" on
 * insert). Editing it here also returns it so the caller can persist
 * the rename together with the format change.
 */
export function FormatChartDialog({ model, seriesNames, onCancel, onConfirm }: Props) {
  const merged = mergeFormat(model);
  const family = CHART_FAMILY_OF[model.type];
  const isAxisFamily = family !== 'pie';
  // Combo (mix bar + line) and a secondary value axis only make sense
  // on a vertical, absolute-scale chart: column / line / area. 100%-
  // stacked variants pin every series to a shared 0–100% scale, and
  // horizontal bars swap the axes so "secondary axis on the right"
  // stops reading like Excel — mirror the gate in `build-option.ts`.
  const is100 = model.type === 'column-stacked-100' || model.type === 'bar-stacked-100';
  const isHorizontalBar = family === 'bar';
  const supportsComboAndDualAxis =
    isAxisFamily && family !== 'scatter' && !is100 && !isHorizontalBar;
  const baseSeriesKind: 'bar' | 'line' = family === 'line' || family === 'area' ? 'line' : 'bar';
  const [title, setTitle] = useState(model.title ?? '');
  const [showTitle, setShowTitle] = useState(merged.showTitle);
  const [legend, setLegend] = useState(merged.legend);
  const [xAxisTitle, setXAxisTitle] = useState(merged.xAxisTitle ?? '');
  const [yAxisTitle, setYAxisTitle] = useState(merged.yAxisTitle ?? '');
  const [gridlines, setGridlines] = useState(merged.gridlines);
  const [dataLabels, setDataLabels] = useState(merged.dataLabels);
  const [palette, setPalette] = useState<ChartPalette>(merged.palette);
  const [trendline, setTrendline] = useState(merged.trendline);
  const [seriesColors, setSeriesColors] = useState<Record<string, string>>(
    merged.seriesColors ?? {},
  );
  const [seriesTypes, setSeriesTypes] = useState<Record<string, 'bar' | 'line'>>(
    merged.seriesTypes ?? {},
  );
  const [secondaryAxis, setSecondaryAxis] = useState<Record<string, boolean>>(
    merged.secondaryAxis ?? {},
  );

  const confirm = () => {
    // Strip empty / palette-matching overrides so the payload only
    // carries explicit user picks.
    const trimmedSeriesColors: Record<string, string> = {};
    for (const [name, color] of Object.entries(seriesColors)) {
      if (color && color.trim()) trimmedSeriesColors[name] = color;
    }
    // Only persist series-kind overrides that actually differ from the
    // chart's base kind, and secondary-axis flags that are `true`, so
    // the payload stays minimal (and toggling back to the default
    // cleanly removes the override).
    const trimmedSeriesTypes: Record<string, 'bar' | 'line'> = {};
    if (supportsComboAndDualAxis) {
      for (const [name, kind] of Object.entries(seriesTypes)) {
        if (kind && kind !== baseSeriesKind) trimmedSeriesTypes[name] = kind;
      }
    }
    const trimmedSecondaryAxis: Record<string, boolean> = {};
    if (supportsComboAndDualAxis) {
      for (const [name, on] of Object.entries(secondaryAxis)) {
        if (on) trimmedSecondaryAxis[name] = true;
      }
    }
    onConfirm({
      showTitle,
      legend,
      ...(xAxisTitle.trim() ? { xAxisTitle: xAxisTitle.trim() } : { xAxisTitle: undefined }),
      ...(yAxisTitle.trim() ? { yAxisTitle: yAxisTitle.trim() } : { yAxisTitle: undefined }),
      gridlines,
      dataLabels,
      palette,
      trendline,
      seriesColors: trimmedSeriesColors,
      seriesTypes: trimmedSeriesTypes,
      secondaryAxis: trimmedSecondaryAxis,
      // Title text is part of the chart's identity (`ChartModel.title`)
      // not its format. The caller reads the trimmed title via the
      // form input below and applies it alongside the format patch.
    });
    // Side-channel the title back via a custom event so the caller
    // doesn't need a separate prop just for one optional string.
    const trimmed = title.trim();
    if (trimmed !== (model.title ?? '')) {
      const ce = new CustomEvent('casual-chart-title-changed', {
        detail: { id: model.id, title: trimmed || undefined },
      });
      document.dispatchEvent(ce);
    }
  };

  return (
    <Dialog
      title="Format chart"
      onClose={onCancel}
      data-testid="format-chart-dialog"
      footer={
        <>
          <button
            type="button"
            className="btn-secondary"
            data-testid="format-chart-cancel"
            onClick={onCancel}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="format-chart-apply"
            onClick={confirm}
          >
            Apply
          </button>
        </>
      }
    >
      <div className="format-chart">
        <Section legend="Title">
          <div className="format-chart__row">
            <label className="format-chart__checkbox">
              <input
                type="checkbox"
                checked={showTitle}
                data-testid="format-chart-show-title"
                onChange={(e) => setShowTitle(e.target.checked)}
              />
              <span>Show title</span>
            </label>
          </div>
          <input
            type="text"
            className="format-chart__input"
            data-testid="format-chart-title-input"
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            placeholder="Chart 1"
            disabled={!showTitle}
          />
        </Section>

        <Section legend="Legend">
          <div className="format-chart__segment" role="radiogroup" aria-label="Legend position">
            {LEGEND_POSITIONS.map((p) => (
              <label
                key={p.id}
                className={`format-chart__seg-opt${legend === p.id ? ' format-chart__seg-opt--active' : ''}`}
                data-testid={`format-chart-legend-${p.id}`}
              >
                <input
                  type="radio"
                  name="legend-pos"
                  value={p.id}
                  checked={legend === p.id}
                  onChange={() => setLegend(p.id)}
                />
                <span>{p.label}</span>
              </label>
            ))}
          </div>
        </Section>

        {isAxisFamily && (
          <Section legend="Axes">
            <div className="format-chart__axis-row">
              <label className="format-chart__field">
                <span className="format-chart__field-label">X-axis title</span>
                <input
                  type="text"
                  className="format-chart__input"
                  data-testid="format-chart-x-axis-title"
                  value={xAxisTitle}
                  onChange={(e) => setXAxisTitle(e.target.value)}
                  placeholder="Auto"
                />
              </label>
              <label className="format-chart__field">
                <span className="format-chart__field-label">Y-axis title</span>
                <input
                  type="text"
                  className="format-chart__input"
                  data-testid="format-chart-y-axis-title"
                  value={yAxisTitle}
                  onChange={(e) => setYAxisTitle(e.target.value)}
                  placeholder="Auto"
                />
              </label>
            </div>
            <label className="format-chart__checkbox">
              <input
                type="checkbox"
                checked={gridlines}
                data-testid="format-chart-gridlines"
                onChange={(e) => setGridlines(e.target.checked)}
              />
              <span>Show major gridlines</span>
            </label>
          </Section>
        )}

        <Section legend="Data labels">
          <label className="format-chart__checkbox">
            <input
              type="checkbox"
              checked={dataLabels}
              data-testid="format-chart-data-labels"
              onChange={(e) => setDataLabels(e.target.checked)}
            />
            <span>Show values on each point</span>
          </label>
        </Section>

        {/* Trendline only meaningful for axis-based charts. Pie /
            doughnut have no time axis, so the toggle is hidden. */}
        {isAxisFamily && (
          <Section legend="Trendline">
            <label className="format-chart__checkbox">
              <input
                type="checkbox"
                checked={trendline}
                data-testid="format-chart-trendline"
                onChange={(e) => setTrendline(e.target.checked)}
              />
              <span>Overlay linear-regression trendline on each series</span>
            </label>
          </Section>
        )}

        {/* Combo + dual axis — per-series render-kind (bar/line) and a
            secondary value axis. Only shown for column / line / area
            charts where it's meaningful, and when the caller supplied
            series names. */}
        {supportsComboAndDualAxis && seriesNames && seriesNames.length > 0 && (
          <Section legend="Series type & axis">
            <p className="format-chart__hint">
              Mix bars and lines, and plot a series against a secondary (right) axis.
            </p>
            <div className="format-chart__series-rows">
              {seriesNames.map((name, idx) => {
                const kind = seriesTypes[name] ?? baseSeriesKind;
                const onSecondary = Boolean(secondaryAxis[name]);
                return (
                  <div
                    key={name}
                    className="format-chart__series-row"
                    data-testid={`format-chart-combo-${idx}`}
                  >
                    <span className="format-chart__series-name">{name}</span>
                    <select
                      className="format-chart__series-kind"
                      data-testid={`format-chart-series-kind-${idx}`}
                      aria-label={`${name} chart type`}
                      value={kind}
                      onChange={(e) =>
                        setSeriesTypes({
                          ...seriesTypes,
                          [name]: e.target.value as 'bar' | 'line',
                        })
                      }
                    >
                      <option value="bar">Bars</option>
                      <option value="line">Line</option>
                    </select>
                    <label className="format-chart__checkbox format-chart__series-secondary">
                      <input
                        type="checkbox"
                        data-testid={`format-chart-secondary-axis-${idx}`}
                        checked={onSecondary}
                        onChange={(e) =>
                          setSecondaryAxis({ ...secondaryAxis, [name]: e.target.checked })
                        }
                      />
                      <span>Secondary axis</span>
                    </label>
                  </div>
                );
              })}
            </div>
          </Section>
        )}

        <Section legend="Colors">
          <div className="format-chart__palettes" role="radiogroup" aria-label="Color palette">
            {(Object.keys(PALETTES) as ChartPalette[]).map((p) => (
              <label
                key={p}
                className={`format-chart__palette${palette === p ? ' format-chart__palette--active' : ''}`}
                data-testid={`format-chart-palette-${p}`}
                title={PALETTE_LABELS[p]}
              >
                <input
                  type="radio"
                  name="palette"
                  value={p}
                  checked={palette === p}
                  onChange={() => setPalette(p)}
                />
                <div className="format-chart__palette-swatches">
                  {PALETTES[p].slice(0, 5).map((c) => (
                    <span
                      key={c}
                      className="format-chart__palette-swatch"
                      style={{ backgroundColor: c }}
                    />
                  ))}
                </div>
                <span className="format-chart__palette-label">{PALETTE_LABELS[p]}</span>
              </label>
            ))}
          </div>
        </Section>

        {/* Per-series colour overrides — only rendered when the caller
            supplied series names from the source range. Each row pairs
            a colour swatch (native picker) with a Reset button that
            removes the override and falls back to the palette. */}
        {seriesNames && seriesNames.length > 0 && (
          <Section legend="Series colors">
            <div className="format-chart__series-rows">
              {seriesNames.map((name, idx) => {
                const override = seriesColors[name] ?? '';
                const defaultColor = PALETTES[palette][idx % PALETTES[palette].length];
                return (
                  <div
                    key={name}
                    className="format-chart__series-row"
                    data-testid={`format-chart-series-${idx}`}
                  >
                    <span className="format-chart__series-name">{name}</span>
                    <input
                      type="color"
                      className="format-chart__series-color"
                      data-testid={`format-chart-series-color-${idx}`}
                      value={override || defaultColor}
                      onChange={(e) => setSeriesColors({ ...seriesColors, [name]: e.target.value })}
                    />
                    {override && (
                      <button
                        type="button"
                        className="format-chart__series-reset"
                        data-testid={`format-chart-series-reset-${idx}`}
                        title="Reset to palette default"
                        onClick={() => {
                          const next = { ...seriesColors };
                          delete next[name];
                          setSeriesColors(next);
                        }}
                      >
                        Reset
                      </button>
                    )}
                  </div>
                );
              })}
            </div>
          </Section>
        )}
      </div>
    </Dialog>
  );
}

function Section({ legend, children }: { legend: string; children: React.ReactNode }) {
  return (
    <fieldset className="format-chart__group">
      <legend className="format-chart__legend">{legend}</legend>
      {children}
    </fieldset>
  );
}
