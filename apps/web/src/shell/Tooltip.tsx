import {
  Children,
  cloneElement,
  isValidElement,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type MouseEvent as ReactMouseEvent,
  type FocusEvent as ReactFocusEvent,
  type ReactElement,
} from 'react';
import { createPortal } from 'react-dom';

type Props = {
  /** Text shown when the user hovers the wrapped element. */
  label: string;
  /** Optional keyboard shortcut, rendered as a dim pill alongside the
   *  label. Excel-style: "Bold   Ctrl+B". When set, the tooltip lays
   *  out as two rows so the shortcut stays scannable. */
  shortcut?: string;
  /** Single trigger element. Its `onMouseEnter` / `onMouseLeave` / `onFocus` /
   * `onBlur` props are merged with the tooltip's own handlers — existing
   * handlers on the trigger still fire. */
  children: ReactElement;
  /** Side relative to the trigger. Defaults to `bottom` (toolbar buttons sit
   * near the top of the chrome, so the tip naturally falls into the grid). */
  side?: 'top' | 'bottom';
  /** ms before the tooltip appears. Defaults to 150 ms — Excel's
   *  toolbar tooltips feel near-instant, and 300 ms read as sluggish
   *  on icon-only buttons where the user is already hunting for the
   *  label. Override per-instance if a slower delay is desired (e.g.
   *  on a frequently-hovered element). */
  delay?: number;
};

/**
 * Lightweight, dependency-free tooltip. Clones its child trigger (no wrapping
 * DOM element — preserves flex layouts in the ribbon) and renders the tip in a
 * portal so it can escape clipping containers. The screen-reader name still
 * comes from the trigger's `aria-label`; the tooltip is `role="tooltip"` and
 * `pointer-events: none` so it never blocks clicks.
 *
 * Reasons to keep this small instead of pulling Radix / Floating UI:
 *   - We only need single-line tooltips on icon buttons; no collision logic.
 *   - Bundle is already heavy with Univer; every kB matters.
 *   - Portals + getBoundingClientRect are stable browser primitives.
 */
export function Tooltip({ label, shortcut, children, side = 'bottom', delay = 150 }: Props) {
  const child = Children.only(children);
  const [coords, setCoords] = useState<{ top: number; left: number } | null>(null);
  // Clamped horizontal centre — the tip is centred on the trigger via
  // translateX(-50%), so for right-edge triggers (zoom controls, panel rail)
  // the raw centre pushes half the tip off-screen. After it renders we measure
  // its width and pull the centre back inside the viewport (8px gutter).
  const [clampedLeft, setClampedLeft] = useState<number | null>(null);
  const tipRef = useRef<HTMLSpanElement>(null);
  const timerRef = useRef<number | null>(null);

  useLayoutEffect(() => {
    if (!coords || !tipRef.current) return;
    const half = tipRef.current.offsetWidth / 2;
    const min = half + 8;
    const max = window.innerWidth - half - 8;
    const next = Math.min(Math.max(coords.left, min), max);
    if (next !== clampedLeft) setClampedLeft(next);
  }, [coords, clampedLeft]);

  // Clean up the pending-show timer on unmount so we don't try to setState on
  // a gone component (e.g. the user clicks a button that conditionally
  // unmounts the trigger).
  useEffect(
    () => () => {
      if (timerRef.current !== null) window.clearTimeout(timerRef.current);
    },
    [],
  );

  if (!isValidElement(child) || !label) return <>{children}</>;

  const place = (el: HTMLElement) => {
    const rect = el.getBoundingClientRect();
    const left = rect.left + rect.width / 2;
    if (side === 'top') {
      setCoords({ top: rect.top - 6, left });
    } else {
      setCoords({ top: rect.bottom + 6, left });
    }
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const childProps = (child as any).props as {
    onMouseEnter?: (e: ReactMouseEvent<HTMLElement>) => void;
    onMouseLeave?: (e: ReactMouseEvent<HTMLElement>) => void;
    onFocus?: (e: ReactFocusEvent<HTMLElement>) => void;
    onBlur?: (e: ReactFocusEvent<HTMLElement>) => void;
  };

  const show = (target: HTMLElement) => {
    if (timerRef.current !== null) window.clearTimeout(timerRef.current);
    timerRef.current = window.setTimeout(() => {
      place(target);
      timerRef.current = null;
    }, delay);
  };

  const hide = () => {
    if (timerRef.current !== null) {
      window.clearTimeout(timerRef.current);
      timerRef.current = null;
    }
    setCoords(null);
    setClampedLeft(null);
  };

  const onMouseEnter = (e: ReactMouseEvent<HTMLElement>) => {
    show(e.currentTarget);
    childProps.onMouseEnter?.(e);
  };
  const onMouseLeave = (e: ReactMouseEvent<HTMLElement>) => {
    hide();
    childProps.onMouseLeave?.(e);
  };
  const onFocus = (e: ReactFocusEvent<HTMLElement>) => {
    show(e.currentTarget);
    childProps.onFocus?.(e);
  };
  const onBlur = (e: ReactFocusEvent<HTMLElement>) => {
    hide();
    childProps.onBlur?.(e);
  };

  const cloned = cloneElement(child, {
    onMouseEnter,
    onMouseLeave,
    onFocus,
    onBlur,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any);

  return (
    <>
      {cloned}
      {coords &&
        createPortal(
          <span
            ref={tipRef}
            role="tooltip"
            data-testid="tooltip"
            className={`tooltip tooltip--${side}${shortcut ? ' tooltip--with-shortcut' : ''}`}
            style={{ top: coords.top, left: clampedLeft ?? coords.left }}
          >
            <span className="tooltip__label">{label}</span>
            {shortcut && (
              <span className="tooltip__shortcut" data-testid="tooltip-shortcut">
                {shortcut}
              </span>
            )}
          </span>,
          document.body,
        )}
    </>
  );
}
