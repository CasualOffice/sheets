import { useEffect, useRef } from 'react';
import { useUniverAPI } from '../use-univer';
import { init, type EChartsType } from './echarts-init';
import { buildEChartsOption } from './build-option';
import type { ChartModel } from './types';

/**
 * Single ECharts instance, positioned via cell-coordinate anchor
 * (`model.pos`). Reads source data on mount and on every Univer
 * value-change event in the source range, redraws the option.
 *
 * P0 — sized + positioned from the parent (ChartLayer hands us
 * `left/top/width/height` already translated to host-local
 * coordinates). P2 hands those off to Univer's drawing model for
 * native move/resize handles; for now the chart is locked in
 * place once inserted.
 */
type Props = {
  model: ChartModel;
  /** Host-local CSS box. Parent computes from the cell rect + the
   *  active sheet's scroll offset on every animation frame. */
  rect: { left: number; top: number; width: number; height: number };
};

export function ChartOverlay({ model, rect }: Props) {
  const api = useUniverAPI();
  const hostRef = useRef<HTMLDivElement>(null);
  const echartRef = useRef<EChartsType | null>(null);

  // Init ECharts once + dispose on unmount. ECharts mutates its
  // own DOM children so it must own its container.
  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const inst = init(host, undefined, { renderer: 'canvas' });
    echartRef.current = inst;
    return () => {
      inst.dispose();
      echartRef.current = null;
    };
  }, []);

  // Resize whenever the cell-rect-derived box changes.
  useEffect(() => {
    echartRef.current?.resize();
  }, [rect.width, rect.height]);

  // Rebuild the option from source data on mount, on model change,
  // and on every Univer value-change. Setting setOption(option,
  // true) replaces (vs. merges) — cleaner when the data shape
  // shifts (e.g. user adds a series column).
  useEffect(() => {
    if (!api) return;
    const refresh = () => {
      const opt = buildEChartsOption(api, model);
      if (opt) echartRef.current?.setOption(opt, true);
    };
    refresh();
    // Watch for cell value changes. We can't easily filter to "only
    // mutations inside the source range" without re-parsing every
    // mutation's params — for v0.1.1 just refresh on any change
    // and trust the cheap setOption(true). Optimization for v0.2.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const evName = (api as any).Event?.SheetValueChanged;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    if (!evName || typeof (api as any).addEvent !== 'function') return;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const disp = (api as any).addEvent(evName, refresh) as { dispose?: () => void };
    return () => disp.dispose?.();
  }, [api, model]);

  return (
    <div
      ref={hostRef}
      className="chart-overlay"
      data-testid="chart-overlay"
      data-chart-id={model.id}
      style={{
        position: 'absolute',
        left: `${rect.left}px`,
        top: `${rect.top}px`,
        width: `${rect.width}px`,
        height: `${rect.height}px`,
      }}
    />
  );
}
