import { useUniverAPI } from '../use-univer';
import { useActiveCellState } from '../hooks/useActiveCellState';
import { Icon } from './Icon';
import {
  decreaseDecimal,
  increaseDecimal,
  setAlignment,
  setNumberFormatByKey,
  toggleBold,
  toggleItalic,
  toggleUnderline,
} from './home-tab-actions';

/**
 * Bottom action bar for mobile (CSS-driven, visible at ≤ 480 px).
 *
 * Real mobile sheet apps put a thumb-reachable formatting strip at the
 * bottom of the screen — the desktop ribbon doesn't translate, and the
 * compact-toolbar-at-top requires the user to crane their thumb up to
 * format a cell they're currently looking at. This component is that
 * strip. Reuses the same `home-tab-actions` commands the desktop ribbon
 * fires, so behaviour and undo grouping are identical.
 *
 * Mounts unconditionally; pure-CSS `display: none` hides it above the
 * mobile breakpoint. No JS-side `window.matchMedia` so the SSR shape
 * and dev-hot-reload shape are the same — no flash on swap.
 */
export function MobileActionBar() {
  const api = useUniverAPI();
  const cell = useActiveCellState();

  if (!api) return null;

  const press = (active: boolean) =>
    active ? 'mobile-bar__btn mobile-bar__btn--on' : 'mobile-bar__btn';

  return (
    <div
      className="mobile-bar"
      role="toolbar"
      aria-label="Formatting"
      data-testid="mobile-action-bar"
    >
      <button
        type="button"
        className={press(cell?.isBold ?? false)}
        aria-pressed={cell?.isBold}
        aria-label="Bold"
        title="Bold"
        onClick={() => toggleBold(api, cell?.isBold ?? false)}
        data-testid="mobile-bar-bold"
      >
        <Icon name="format_bold" />
      </button>
      <button
        type="button"
        className={press(cell?.isItalic ?? false)}
        aria-pressed={cell?.isItalic}
        aria-label="Italic"
        title="Italic"
        onClick={() => toggleItalic(api, cell?.isItalic ?? false)}
        data-testid="mobile-bar-italic"
      >
        <Icon name="format_italic" />
      </button>
      <button
        type="button"
        className={press(cell?.isUnderline ?? false)}
        aria-pressed={cell?.isUnderline}
        aria-label="Underline"
        title="Underline"
        onClick={() => toggleUnderline(api, cell?.isUnderline ?? false)}
        data-testid="mobile-bar-underline"
      >
        <Icon name="format_underlined" />
      </button>

      <span className="mobile-bar__sep" aria-hidden="true" />

      <button
        type="button"
        className="mobile-bar__btn"
        aria-label="Format as currency"
        title="Currency"
        onClick={() => setNumberFormatByKey(api, 'currency')}
        data-testid="mobile-bar-currency"
      >
        <Icon name="attach_money" />
      </button>
      <button
        type="button"
        className="mobile-bar__btn"
        aria-label="Format as percent"
        title="Percent"
        onClick={() => setNumberFormatByKey(api, 'percent')}
        data-testid="mobile-bar-percent"
      >
        <Icon name="percent" />
      </button>
      <button
        type="button"
        className="mobile-bar__btn"
        aria-label="Decrease decimal places"
        title="Decrease decimal"
        onClick={() => decreaseDecimal(api)}
        data-testid="mobile-bar-dec-decimal"
      >
        <Icon name="decimal_decrease" />
      </button>
      <button
        type="button"
        className="mobile-bar__btn"
        aria-label="Increase decimal places"
        title="Increase decimal"
        onClick={() => increaseDecimal(api)}
        data-testid="mobile-bar-inc-decimal"
      >
        <Icon name="decimal_increase" />
      </button>

      <span className="mobile-bar__sep" aria-hidden="true" />

      <button
        type="button"
        className={
          cell?.align === 'left'
            ? 'mobile-bar__btn mobile-bar__btn--on'
            : 'mobile-bar__btn'
        }
        aria-pressed={cell?.align === 'left'}
        aria-label="Align left"
        title="Align left"
        onClick={() => setAlignment(api, 'left')}
        data-testid="mobile-bar-align-left"
      >
        <Icon name="format_align_left" />
      </button>
      <button
        type="button"
        className={
          cell?.align === 'center'
            ? 'mobile-bar__btn mobile-bar__btn--on'
            : 'mobile-bar__btn'
        }
        aria-pressed={cell?.align === 'center'}
        aria-label="Align center"
        title="Align center"
        onClick={() => setAlignment(api, 'center')}
        data-testid="mobile-bar-align-center"
      >
        <Icon name="format_align_center" />
      </button>
      <button
        type="button"
        className={
          cell?.align === 'right'
            ? 'mobile-bar__btn mobile-bar__btn--on'
            : 'mobile-bar__btn'
        }
        aria-pressed={cell?.align === 'right'}
        aria-label="Align right"
        title="Align right"
        onClick={() => setAlignment(api, 'right')}
        data-testid="mobile-bar-align-right"
      >
        <Icon name="format_align_right" />
      </button>
    </div>
  );
}
