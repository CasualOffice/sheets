import {
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type ReactNode,
  type RefObject,
} from 'react';
import { createPortal } from 'react-dom';

/**
 * Anchored popover. Positions itself just below the anchor element, closes on
 * outside click and Escape. Reused for dropdown menus (File menu, Borders,
 * future Number Format, etc.).
 */
type Props = {
  anchorRef: RefObject<HTMLElement | null>;
  onClose: () => void;
  className?: string;
  ['data-testid']?: string;
  children: ReactNode;
};

export function Popover({ anchorRef, onClose, className, children, ...rest }: Props) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ top: number; left: number }>({ top: 0, left: 0 });

  useLayoutEffect(() => {
    const rect = anchorRef.current?.getBoundingClientRect();
    if (rect) setPos({ top: rect.bottom + 2, left: rect.left });
  }, [anchorRef]);

  useEffect(() => {
    const onDocPointerDown = (e: PointerEvent) => {
      if (
        !ref.current?.contains(e.target as Node) &&
        !anchorRef.current?.contains(e.target as Node)
      ) {
        onClose();
      }
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    // Capture phase + pointerdown — Univer's canvas calls stopPropagation
    // on mousedown, so a bubble-phase document listener never fires for grid
    // clicks. Capturing pointerdown gives us the event before any handler
    // can swallow it.
    document.addEventListener('pointerdown', onDocPointerDown, true);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('pointerdown', onDocPointerDown, true);
      document.removeEventListener('keydown', onKey);
    };
  }, [onClose, anchorRef]);

  // Portaled to <body> so the popover escapes every parent stacking context
  // (toolbar, ribbon group, btn-split, etc.). Without this, even with a
  // high z-index, the rendered popover was sitting below the formula bar
  // input — Playwright caught it because pointer hit-tests picked the input.
  return createPortal(
    <div
      ref={ref}
      className={`menu ${className ?? ''}`}
      style={{ top: pos.top, left: pos.left }}
      role="menu"
      data-testid={rest['data-testid']}
    >
      {children}
    </div>,
    document.body,
  );
}
