import { useRef, useState, type ChangeEvent, type ReactNode } from 'react';
import { Icon } from './Icon';
import { Popover } from './Popover';
import { Tooltip } from './Tooltip';

/** Reusable ribbon primitives shared across all tabs. */

export function RibbonGroup({
  label,
  rows,
  children,
}: {
  label: string;
  /** Stack children vertically as separate rows. */
  rows?: boolean;
  children: ReactNode;
}) {
  return (
    <div className="ribbon__group" data-testid={`ribbon-group-${label.toLowerCase()}`}>
      <div className={`ribbon__group-body${rows ? ' ribbon__group-body--rows' : ''}`}>
        {children}
      </div>
      <div className="ribbon__group-label">{label}</div>
    </div>
  );
}

export function RibbonRow({ children }: { children: ReactNode }) {
  return <div className="ribbon__row">{children}</div>;
}

type ToolbarButtonProps = {
  id: string;
  label: string;
  icon: string;
  pressed?: boolean;
  disabled?: boolean;
  onClick?: () => void;
};

export function ToolbarButton({
  id,
  label,
  icon,
  pressed,
  disabled,
  onClick,
}: ToolbarButtonProps) {
  return (
    <Tooltip label={label}>
      <button
        type="button"
        className="btn btn--icon"
        data-testid={`ribbon-btn-${id}`}
        aria-label={label}
        aria-pressed={pressed}
        disabled={disabled}
        onClick={onClick}
      >
        <Icon name={icon} size="sm" />
      </button>
    </Tooltip>
  );
}

type SelectProps = {
  id: string;
  label: string;
  value: string;
  options: { value: string; label: string }[];
  width?: number;
  disabled?: boolean;
  onChange: (v: string) => void;
};

export function ToolbarSelect({
  id,
  label,
  value,
  options,
  width = 110,
  disabled,
  onChange,
}: SelectProps) {
  return (
    <Tooltip label={label}>
      <select
        data-testid={`ribbon-select-${id}`}
        aria-label={label}
        disabled={disabled}
        value={value}
        style={{ width }}
        className="ribbon__select"
        onChange={(e: ChangeEvent<HTMLSelectElement>) => onChange(e.target.value)}
      >
        {options.map((o) => (
          <option key={o.value} value={o.value}>
            {o.label}
          </option>
        ))}
      </select>
    </Tooltip>
  );
}

type ColorButtonProps = {
  id: string;
  label: string;
  icon: string;
  /** Current color in hex (`#RRGGBB`) or empty. */
  value: string;
  /** Color to apply when the icon (not the swatch) is clicked. */
  defaultColor: string;
  disabled?: boolean;
  onChange: (color: string) => void;
};

/**
 * Split button: icon applies the default action; caret opens a popover.
 * Used for Borders and any future "icon + dropdown options" toolbar control.
 */
type DropdownItem = { id: string; label: string; icon: string };

export function ToolbarDropdown({
  id,
  label,
  icon,
  items,
  disabled,
  onChoose,
  onDefault,
}: {
  id: string;
  label: string;
  icon: string;
  items: DropdownItem[];
  disabled?: boolean;
  /** Called when an item is picked. */
  onChoose: (itemId: string) => void;
  /** Called when the icon (not the caret) is clicked — applies the last/default. */
  onDefault?: () => void;
}) {
  const anchorRef = useRef<HTMLDivElement>(null);
  const [open, setOpen] = useState(false);

  return (
    <span
      ref={anchorRef}
      className="btn-split"
      data-testid={`ribbon-dropdown-${id}`}
    >
      <Tooltip label={label}>
        <button
          type="button"
          className="btn btn--icon btn-split__icon"
          data-testid={`ribbon-dropdown-${id}-apply`}
          aria-label={label}
          disabled={disabled}
          onClick={() => {
            onDefault?.();
            // Office pattern: applying the icon's default also dismisses the
            // popover if it happens to be open.
            if (open) setOpen(false);
          }}
        >
          <Icon name={icon} size="sm" />
        </button>
      </Tooltip>
      <button
        type="button"
        className="btn btn--icon btn-split__caret"
        data-testid={`ribbon-dropdown-${id}-caret`}
        aria-haspopup="menu"
        aria-expanded={open}
        aria-label={`${label} options`}
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
      >
        <Icon name="arrow_drop_down" size="sm" />
      </button>
      {open && (
        <Popover
          anchorRef={anchorRef}
          onClose={() => setOpen(false)}
          data-testid={`ribbon-dropdown-${id}-popover`}
        >
          {items.map((item) => (
            <button
              key={item.id}
              type="button"
              className="menu__item"
              role="menuitem"
              data-testid={`ribbon-dropdown-${id}-item-${item.id}`}
              onClick={() => {
                onChoose(item.id);
                setOpen(false);
              }}
            >
              <Icon name={item.icon} size="sm" className="menu__item-icon" />
              <span>{item.label}</span>
            </button>
          ))}
        </Popover>
      )}
    </span>
  );
}

/**
 * Borders split-button: like ToolbarDropdown, but the popover also has a
 * color-picker row at the bottom so the user can change the border color
 * before applying. The selected color persists for the lifetime of the
 * component (one workbook session) — matches Excel's "sticks until you
 * change it" behavior.
 */
const BORDER_COLOR_PRESETS = [
  '#000000', '#666666', '#9aa0a6', '#d2d6dc',
  '#d93025', '#e8710a', '#f9ab00', '#188038',
  '#1a73e8', '#7627bb', '#a142f4', '#e91e63',
];

export function BordersControl({
  label,
  icon,
  items,
  disabled,
  defaultColor,
  onChoose,
  onDefault,
}: {
  label: string;
  icon: string;
  items: DropdownItem[];
  disabled?: boolean;
  /** Initial color before the user picks one. */
  defaultColor: string;
  /** Fires with both the chosen border style and the currently-selected color. */
  onChoose: (itemId: string, color: string) => void;
  /** Fires when the icon (not the caret) is clicked. Receives the current color. */
  onDefault?: (color: string) => void;
}) {
  const anchorRef = useRef<HTMLDivElement>(null);
  const [open, setOpen] = useState(false);
  const [color, setColor] = useState(defaultColor);

  return (
    <span
      ref={anchorRef}
      className="btn-split"
      data-testid="ribbon-dropdown-borders"
    >
      <Tooltip label={`${label} (${color})`}>
        <button
          type="button"
          className="btn btn--icon btn-split__icon btn-color__icon"
          data-testid="ribbon-dropdown-borders-apply"
          aria-label={label}
          disabled={disabled}
          onClick={() => {
            onDefault?.(color);
            if (open) setOpen(false);
          }}
        >
          <Icon name={icon} size="sm" />
          <span
            className="btn-color__swatch"
            style={{ background: color }}
            aria-hidden="true"
          />
        </button>
      </Tooltip>
      <button
        type="button"
        className="btn btn--icon btn-split__caret"
        data-testid="ribbon-dropdown-borders-caret"
        aria-haspopup="menu"
        aria-expanded={open}
        aria-label={`${label} options`}
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
      >
        <Icon name="arrow_drop_down" size="sm" />
      </button>
      {open && (
        <Popover
          anchorRef={anchorRef}
          onClose={() => setOpen(false)}
          data-testid="ribbon-dropdown-borders-popover"
        >
          {items.map((item) => (
            <button
              key={item.id}
              type="button"
              className="menu__item"
              role="menuitem"
              data-testid={`ribbon-dropdown-borders-item-${item.id}`}
              onClick={() => {
                onChoose(item.id, color);
                setOpen(false);
              }}
            >
              <Icon name={item.icon} size="sm" className="menu__item-icon" />
              <span>{item.label}</span>
            </button>
          ))}
          <div className="menu__divider" />
          <div className="menu__color-row" data-testid="ribbon-dropdown-borders-color-row">
            <div className="menu__color-label">Line color</div>
            <div className="menu__color-swatches" role="group" aria-label="Border line color">
              {BORDER_COLOR_PRESETS.map((preset) => (
                <button
                  key={preset}
                  type="button"
                  className={`menu__color-swatch${color === preset ? ' menu__color-swatch--active' : ''}`}
                  style={{ background: preset }}
                  data-testid={`ribbon-dropdown-borders-color-${preset.slice(1)}`}
                  aria-label={`Set border color to ${preset}`}
                  aria-pressed={color === preset}
                  onClick={() => setColor(preset)}
                />
              ))}
              <label
                className="menu__color-swatch menu__color-swatch--custom"
                title="Custom color"
              >
                <Icon name="palette" size="sm" />
                <input
                  type="color"
                  value={color}
                  data-testid="ribbon-dropdown-borders-color-custom"
                  onChange={(e) => setColor(e.target.value)}
                />
              </label>
            </div>
          </div>
        </Popover>
      )}
    </span>
  );
}

/**
 * Combo: clicking the icon applies the last-chosen color; clicking the small
 * dropdown caret opens a native color picker. Mirrors Excel's split button.
 */
export function ToolbarColorButton({
  id,
  label,
  icon,
  value,
  defaultColor,
  disabled,
  onChange,
}: ColorButtonProps) {
  const applied = value || defaultColor;
  return (
    <span className="btn-color" data-testid={`ribbon-color-${id}`}>
      <Tooltip label={label}>
        <button
          type="button"
          className="btn btn--icon btn-color__icon"
          data-testid={`ribbon-color-${id}-apply`}
          aria-label={label}
          disabled={disabled}
          onClick={() => onChange(applied)}
        >
          <Icon name={icon} size="sm" />
          <span className="btn-color__swatch" style={{ background: applied }} aria-hidden="true" />
        </button>
      </Tooltip>
      <label className="btn-color__picker" title={`${label} — choose color`}>
        <Icon name="arrow_drop_down" size="sm" />
        <input
          type="color"
          data-testid={`ribbon-color-${id}-input`}
          value={applied || '#000000'}
          disabled={disabled}
          onChange={(e: ChangeEvent<HTMLInputElement>) => onChange(e.target.value)}
          tabIndex={-1}
        />
      </label>
    </span>
  );
}
