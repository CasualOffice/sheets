import type { CSSProperties } from 'react';

export interface SelectionStats {
  sum?: number | null;
  avg?: number | null;
  numCount?: number | null;
  count?: number | null;
}

export interface StatusBarProps {
  /** Live presence indicator on the left (e.g. "Solo", "In room", "Offline"). */
  mode?: string;
  /** Stats for the active selection. */
  stats?: SelectionStats;
  /** Current zoom percentage (0-200). */
  zoom?: number;
  onZoomChange?: (zoom: number) => void;
  /** Hide the auxiliary stats (avg / num / zoom slider) on narrow widths. */
  compact?: boolean;
  style?: CSSProperties;
}

const formatNumber = (n: number): string =>
  n.toLocaleString('en-US', { maximumFractionDigits: 2 });

export function StatusBar({
  mode = 'Solo',
  stats = {},
  zoom = 100,
  onZoomChange,
  compact = false,
  style,
}: StatusBarProps) {
  return (
    <div
      style={{
        height: 'var(--statusbar-h)',
        flex: '0 0 var(--statusbar-h)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        padding: '0 16px',
        background: 'var(--color-surface)',
        borderTop: '1px solid var(--color-divider)',
        fontFamily: 'var(--font-sans)',
        fontSize: 'var(--text-xs)',
        color: 'var(--color-text-muted)',
        ...style,
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
        <span
          style={{
            display: 'inline-flex',
            alignItems: 'center',
            gap: 5,
            color: 'var(--color-success)',
          }}
        >
          <span
            style={{
              width: 6,
              height: 6,
              borderRadius: '50%',
              background: 'var(--color-success)',
            }}
          />
          {mode}
        </span>
        {!compact && <span>Ready</span>}
      </div>
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: 16,
          fontVariantNumeric: 'tabular-nums',
        }}
      >
        {stats.sum != null && (
          <Stat label="Sum" value={formatNumber(stats.sum)} />
        )}
        {!compact && stats.avg != null && (
          <Stat label="Avg" value={formatNumber(stats.avg)} />
        )}
        {!compact && stats.numCount != null && (
          <Stat label="Num" value={String(stats.numCount)} />
        )}
        {stats.count != null && (
          <Stat label="Count" value={String(stats.count)} />
        )}
        {!compact && (
          <ZoomSlider zoom={zoom} onChange={onZoomChange} />
        )}
      </div>
    </div>
  );
}

function Stat({ label, value }: { label: string; value: string }) {
  return (
    <span>
      {label}{' '}
      <strong style={{ color: 'var(--color-text)', fontWeight: 'var(--weight-semibold)' }}>
        {value}
      </strong>
    </span>
  );
}

function ZoomSlider({
  zoom,
  onChange,
}: {
  zoom: number;
  onChange?: (next: number) => void;
}) {
  const pct = Math.max(0, Math.min(200, zoom));
  const left = (pct / 200) * 100;
  return (
    <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, cursor: onChange ? 'pointer' : 'default' }}>
      <span
        style={{
          width: 64,
          height: 4,
          background: 'var(--color-surface-alt)',
          borderRadius: 999,
          position: 'relative',
        }}
      >
        <span
          style={{
            position: 'absolute',
            left: `${left}%`,
            top: -2,
            width: 8,
            height: 8,
            borderRadius: '50%',
            background: 'var(--color-accent)',
            transform: 'translateX(-50%)',
          }}
        />
        {onChange && (
          <input
            type="range"
            min={10}
            max={200}
            step={5}
            value={pct}
            onChange={(e) => onChange(Number(e.currentTarget.value))}
            style={{
              position: 'absolute',
              inset: -4,
              opacity: 0,
              cursor: 'pointer',
            }}
          />
        )}
      </span>
      {pct}%
    </label>
  );
}
