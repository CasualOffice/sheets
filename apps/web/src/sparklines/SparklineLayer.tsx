import { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { useUniverAPI } from '../use-univer';
import { getHeaderGutter, getUniverHost, getUniverMainCanvas } from '../univer-dom';
import { useSparklines } from './sparklines-context';
import type { SparklineModel } from './types';

/**
 * Renders an in-cell mini-chart per sparkline definition. Mirrors
 * ChartLayer's positioning (cell-anchored via Univer's `getCellRect`,
 * rAF-driven so scroll / zoom track) but the geometry is small enough
 * that we render SVG directly instead of mounting ECharts per
 * sparkline — that would be ~30 ms per init, way too slow for a
 * column of 50 sparklines.
 *
 * Three sparkline types per Excel:
 *   - **line**: polyline through the values; min / max markers.
 *   - **column**: short bars per value, positive in accent colour,
 *     negative tinted.
 *   - **win-loss**: tri-state bars (+1 / -1 / 0).
 *
 * Sparkline data is read fresh every frame so cell edits to the
 * source range update the visual immediately — for normal sizes (a
 * few dozen sparklines, source ranges of < 50 cells each) this is
 * sub-millisecond per frame.
 */

const PAD = 2; // px inset from cell edges

type Geom = {
  sparkline: SparklineModel;
  values: Array<number | null>;
  left: number;
  top: number;
  width: number;
  height: number;
};

export function SparklineLayer() {
  const api = useUniverAPI();
  const { sparklines } = useSparklines();
  const [geoms, setGeoms] = useState<Geom[]>([]);
  const hostRef = useRef<HTMLElement | null>(null);
  const geomsRef = useRef<Geom[]>(geoms);
  geomsRef.current = geoms;

  useEffect(() => {
    hostRef.current = getUniverHost();
  }, []);

  useEffect(() => {
    if (!api) return;
    if (sparklines.length === 0) {
      if (geomsRef.current.length) setGeoms([]);
      return;
    }
    let raf = 0;
    const tick = () => {
      try {
        recompute();
      } catch (err) {
        console.debug('[sparklines] recompute threw', err);
      }
      raf = requestAnimationFrame(tick);
    };
    const recompute = () => {
      const host = hostRef.current ?? getUniverHost();
      if (!host) {
        if (geomsRef.current.length) setGeoms([]);
        return;
      }
      const canvas = getUniverMainCanvas(host);
      if (!canvas) return;
      const hostRect = host.getBoundingClientRect();
      const canvasRect = canvas.getBoundingClientRect();
      const dx = canvasRect.left - hostRect.left;
      const dy = canvasRect.top - hostRect.top;
      const gutter = getHeaderGutter(api);

      const wb = api.getActiveWorkbook();
      if (!wb) return;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const activeSheet = wb.getActiveSheet() as any;
      const activeSheetId = activeSheet?.getSheetId?.() ?? '';

      let sx = 0;
      let sy = 0;
      try {
        const scrollState = activeSheet?.getScrollState?.() as
          | { sheetViewStartRow?: number; sheetViewStartColumn?: number; offsetX?: number; offsetY?: number }
          | undefined;
        if (scrollState) {
          const r = scrollState.sheetViewStartRow ?? 0;
          const c = scrollState.sheetViewStartColumn ?? 0;
          const topLeft = activeSheet.getRange(r, c).getCellRect();
          if (topLeft) {
            sx = topLeft.left + (scrollState.offsetX ?? 0);
            sy = topLeft.top + (scrollState.offsetY ?? 0);
          }
        }
      } catch {
        /* skeleton not ready */
      }

      let zoom = 1;
      try {
        const z =
          (activeSheet?._worksheet?.getZoomRatio?.() as number | undefined) ??
          (activeSheet?.getZoomRatio?.() as number | undefined);
        if (typeof z === 'number' && z > 0) zoom = z;
      } catch {
        /* default 1 */
      }

      const out: Geom[] = [];
      for (const s of sparklines) {
        // Cross-sheet sparklines: render only when the anchor sheet
        // matches the active sheet. The data range is always read
        // from the sparkline's own sheetId so values stay
        // consistent even if the user has scrolled to a different
        // sheet (which then wouldn't render the overlay).
        if (s.sheetId !== activeSheetId) continue;
        try {
          const rect = activeSheet.getRange(s.anchor.row, s.anchor.col).getCellRect();
          if (!rect) continue;
          const left = (rect.left - sx) * zoom + dx + gutter.rowHeaderWidth + PAD;
          const top = (rect.top - sy) * zoom + dy + gutter.columnHeaderHeight + PAD;
          const right = (rect.right - sx) * zoom + dx + gutter.rowHeaderWidth - PAD;
          const bottom = (rect.bottom - sy) * zoom + dy + gutter.columnHeaderHeight - PAD;
          if (right < dx || bottom < dy) continue;
          if (left > dx + canvasRect.width || top > dy + canvasRect.height) continue;
          // Read source values for this sparkline.
          const values: Array<number | null> = [];
          for (let r = s.source.startRow; r <= s.source.endRow; r += 1) {
            for (let c = s.source.startColumn; c <= s.source.endColumn; c += 1) {
              const v = activeSheet.getRange(r, c).getValue();
              const n = typeof v === 'number' ? v : Number(v);
              values.push(Number.isFinite(n) ? n : null);
            }
          }
          out.push({
            sparkline: s,
            values,
            left,
            top,
            width: Math.max(8, right - left),
            height: Math.max(8, bottom - top),
          });
        } catch {
          /* getCellRect mid-resize / source out-of-bounds */
        }
      }
      if (geomEqual(out, geomsRef.current)) return;
      setGeoms(out);
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
  }, [api, sparklines]);

  const host = hostRef.current ?? getUniverHost();
  if (!host) return null;
  if (geoms.length === 0) return null;

  return createPortal(
    <div className="sparkline-layer" data-testid="sparkline-layer" aria-hidden="true">
      {geoms.map((g) => (
        <div
          key={g.sparkline.id}
          className="sparkline-cell"
          data-testid="sparkline-cell"
          style={{ left: g.left, top: g.top, width: g.width, height: g.height }}
        >
          <SparklineSvg geom={g} />
        </div>
      ))}
    </div>,
    host,
  );
}

function SparklineSvg({ geom }: { geom: Geom }) {
  const { sparkline, values, width, height } = geom;
  const color = sparkline.color ?? '#0c825a';
  const neg = sparkline.negativeColor ?? '#d93025';
  // Numeric-only series for the value-axis range. Null cells act as
  // gaps in line / column modes.
  const nums = values.filter((v): v is number => typeof v === 'number');
  if (nums.length === 0) return null;
  const min = Math.min(...nums);
  const max = Math.max(...nums);
  const range = max - min || 1;
  const n = values.length;

  if (sparkline.type === 'line') {
    const points: string[] = [];
    let minIdx = -1;
    let maxIdx = -1;
    for (let i = 0; i < n; i += 1) {
      const v = values[i];
      if (v == null) continue;
      const x = n > 1 ? (i / (n - 1)) * width : width / 2;
      const y = height - ((v - min) / range) * height;
      points.push(`${x},${y}`);
      if (v === min) minIdx = i;
      if (v === max) maxIdx = i;
    }
    const pt = (idx: number): { x: number; y: number } | null => {
      const v = values[idx];
      if (v == null) return null;
      const x = n > 1 ? (idx / (n - 1)) * width : width / 2;
      const y = height - ((v - min) / range) * height;
      return { x, y };
    };
    const minP = minIdx >= 0 ? pt(minIdx) : null;
    const maxP = maxIdx >= 0 ? pt(maxIdx) : null;
    return (
      <svg
        width={width}
        height={height}
        viewBox={`0 0 ${width} ${height}`}
        xmlns="http://www.w3.org/2000/svg"
      >
        <polyline points={points.join(' ')} fill="none" stroke={color} strokeWidth={1.25} strokeLinejoin="round" strokeLinecap="round" />
        {maxP && <circle cx={maxP.x} cy={maxP.y} r={1.6} fill={color} />}
        {minP && <circle cx={minP.x} cy={minP.y} r={1.6} fill={neg} />}
      </svg>
    );
  }

  if (sparkline.type === 'column') {
    const gap = 1;
    const barWidth = Math.max(1, (width - gap * (n - 1)) / n);
    // Bars sit on the baseline = where 0 lives in the value range, or
    // the chart bottom if all values are ≥ 0.
    const zeroY = max <= 0 ? 0 : min >= 0 ? height : height - ((0 - min) / range) * height;
    return (
      <svg
        width={width}
        height={height}
        viewBox={`0 0 ${width} ${height}`}
        xmlns="http://www.w3.org/2000/svg"
      >
        {values.map((v, i) => {
          if (v == null) return null;
          const x = i * (barWidth + gap);
          const y = height - ((v - min) / range) * height;
          const yTop = Math.min(y, zeroY);
          const h = Math.max(0.5, Math.abs(y - zeroY));
          return (
            <rect
              key={i}
              x={x}
              y={yTop}
              width={barWidth}
              height={h}
              fill={v < 0 ? neg : color}
            />
          );
        })}
      </svg>
    );
  }

  // win-loss
  const gap = 1;
  const barWidth = Math.max(1, (width - gap * (n - 1)) / n);
  const halfH = height / 2;
  return (
    <svg
      width={width}
      height={height}
      viewBox={`0 0 ${width} ${height}`}
      xmlns="http://www.w3.org/2000/svg"
    >
      {values.map((v, i) => {
        if (v == null || v === 0) return null;
        const x = i * (barWidth + gap);
        return (
          <rect
            key={i}
            x={x}
            y={v > 0 ? 0 : halfH}
            width={barWidth}
            height={halfH}
            fill={v > 0 ? color : neg}
          />
        );
      })}
    </svg>
  );
}

function geomEqual(a: Geom[], b: Geom[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) {
    const x = a[i];
    const y = b[i];
    if (
      x.sparkline.id !== y.sparkline.id ||
      x.left !== y.left ||
      x.top !== y.top ||
      x.width !== y.width ||
      x.height !== y.height ||
      x.values.length !== y.values.length
    )
      return false;
    for (let j = 0; j < x.values.length; j += 1) {
      if (x.values[j] !== y.values[j]) return false;
    }
  }
  return true;
}
