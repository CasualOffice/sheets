import { useEffect, useRef, type ReactNode } from 'react';
import { Icon } from './Icon';

type Props = {
  title: string;
  onClose: () => void;
  footer?: ReactNode;
  children: ReactNode;
  ['data-testid']?: string;
};

/**
 * Generic modal dialog. Click-backdrop and Escape close it. Focus is trapped
 * minimally — first focusable element gets focus on open.
 */
export function Dialog({ title, onClose, footer, children, ...rest }: Props) {
  const dialogRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    document.addEventListener('keydown', onKey);

    const first = dialogRef.current?.querySelector<HTMLElement>(
      'input, textarea, button, [tabindex]:not([tabindex="-1"])',
    );
    first?.focus();

    return () => document.removeEventListener('keydown', onKey);
  }, [onClose]);

  return (
    <div
      className="dialog-backdrop"
      data-testid="dialog-backdrop"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        className="dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="dialog-title"
        ref={dialogRef}
        data-testid={rest['data-testid']}
      >
        <div className="dialog__header">
          <h2 className="dialog__title" id="dialog-title">
            {title}
          </h2>
          <button
            type="button"
            className="btn btn--icon"
            data-testid="dialog-close"
            aria-label="Close"
            onClick={onClose}
          >
            <Icon name="close" size="sm" />
          </button>
        </div>
        <div className="dialog__body">{children}</div>
        {footer && <div className="dialog__footer">{footer}</div>}
      </div>
    </div>
  );
}
