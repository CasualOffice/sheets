import { useActiveCellState } from '../hooks/useActiveCellState';

const NUM = new Intl.NumberFormat(undefined, { maximumFractionDigits: 4 });

export function StatusBar() {
  const { stats, ready } = useActiveCellState();

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
      <span data-testid="statusbar-zoom">100%</span>
    </footer>
  );
}
