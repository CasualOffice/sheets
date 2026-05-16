import { useEffect, useState } from 'react';
import { useActiveCellState } from '../hooks/useActiveCellState';
import { useUniverAPI } from '../use-univer';
import { setZoom } from './tab-actions';

const NUM = new Intl.NumberFormat(undefined, { maximumFractionDigits: 4 });

const ZOOM_STEPS = [25, 50, 75, 100, 125, 150, 200, 300, 400];

export function StatusBar() {
  const api = useUniverAPI();
  const { stats, ready } = useActiveCellState();
  const [zoomPct, setZoomPct] = useState(100);

  // Sync zoomPct with whatever the workbook's actual zoom ratio is. We listen
  // to the set-zoom-ratio command (which carries the new value in its params)
  // — that covers BOTH our ribbon button paths AND any future paths that
  // dispatch the same command.
  useEffect(() => {
    if (!api) return;
    const d = api.addEvent(api.Event.CommandExecuted, (e) => {
      const info = e as { id?: string; params?: { zoomRatio?: number } };
      if (info.id === 'sheet.command.set-zoom-ratio' && typeof info.params?.zoomRatio === 'number') {
        setZoomPct(Math.round(info.params.zoomRatio * 100));
      }
    });
    return () => d.dispose();
  }, [api]);

  const applyZoom = (pct: number) => {
    if (!api) return;
    const clamped = Math.max(25, Math.min(400, pct));
    setZoomPct(clamped);
    setZoom(api, clamped / 100);
  };

  return (
    <footer className="statusbar" data-testid="statusbar" role="status">
      <span>
        <span className="statusbar__dot" aria-hidden="true" />
        {ready ? 'Ready' : 'Loading…'}
      </span>
      <span className="statusbar__spacer" />
      {stats && stats.count > 0 && (
        <span className="statusbar__stats" data-testid="statusbar-stats">
          <span data-testid="stat-count">Count: {stats.count}</span>
          <span data-testid="stat-sum">Sum: {NUM.format(stats.sum)}</span>
          {stats.avg !== null && (
            <span data-testid="stat-avg">Avg: {NUM.format(stats.avg)}</span>
          )}
        </span>
      )}
      <div className="statusbar__zoom" role="group" aria-label="Zoom">
        <button
          type="button"
          className="statusbar__zoom-btn"
          data-testid="statusbar-zoom-out"
          aria-label="Zoom out"
          onClick={() => {
            const prev = [...ZOOM_STEPS].reverse().find((s) => s < zoomPct);
            applyZoom(prev ?? 25);
          }}
        >
          −
        </button>
        <input
          type="range"
          min={25}
          max={400}
          step={5}
          value={zoomPct}
          data-testid="statusbar-zoom-slider"
          aria-label="Zoom slider"
          onChange={(e) => applyZoom(Number(e.target.value))}
        />
        <button
          type="button"
          className="statusbar__zoom-btn"
          data-testid="statusbar-zoom-in"
          aria-label="Zoom in"
          onClick={() => {
            const next = ZOOM_STEPS.find((s) => s > zoomPct);
            applyZoom(next ?? 400);
          }}
        >
          +
        </button>
        <button
          type="button"
          className="statusbar__zoom-label-btn"
          data-testid="statusbar-zoom-label"
          aria-label="Reset zoom to 100%"
          title="Reset to 100%"
          onClick={() => applyZoom(100)}
        >
          {zoomPct}%
        </button>
      </div>
    </footer>
  );
}
