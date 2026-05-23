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
export function FormatChartDialog({ model, onCancel, onConfirm }: Props) {
  const merged = mergeFormat(model);
  const family = CHART_FAMILY_OF[model.type];
  const isAxisFamily = family !== 'pie';
  const [title, setTitle] = useState(model.title ?? '');
  const [showTitle, setShowTitle] = useState(merged.showTitle);
  const [legend, setLegend] = useState(merged.legend);
  const [xAxisTitle, setXAxisTitle] = useState(merged.xAxisTitle ?? '');
  const [yAxisTitle, setYAxisTitle] = useState(merged.yAxisTitle ?? '');
  const [gridlines, setGridlines] = useState(merged.gridlines);
  const [dataLabels, setDataLabels] = useState(merged.dataLabels);
  const [palette, setPalette] = useState<ChartPalette>(merged.palette);
  const [trendline, setTrendline] = useState(merged.trendline);

  const confirm = () => {
    onConfirm({
      showTitle,
      legend,
      ...(xAxisTitle.trim() ? { xAxisTitle: xAxisTitle.trim() } : { xAxisTitle: undefined }),
      ...(yAxisTitle.trim() ? { yAxisTitle: yAxisTitle.trim() } : { yAxisTitle: undefined }),
      gridlines,
      dataLabels,
      palette,
      trendline,
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
