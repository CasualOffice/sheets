/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { useRef, useState, type ChangeEvent, type ReactNode } from 'react';
import { Icon } from './Icon';
import { Popover } from './Popover';
import { Tooltip } from './Tooltip';

/** Reusable ribbon primitives shared across all tabs. */

export function RibbonGroup({
  label,
  rows,
  lead,
  row1,
  row2,
  children,
}: {
  label: string;
  /** Stack children vertically as separate rows (legacy single-children API). */
  rows?: boolean;
  /** Optional "primary" big button on the left of the group that spans
   *  both rows. Use BigToolbarButton here. Excel uses this slot for
   *  group-defining actions: Paste in Clipboard, Styles button, etc. */
  lead?: ReactNode;
  /** Top row of small buttons (high-frequency actions). */
  row1?: ReactNode;
  /** Bottom row of small buttons (secondary actions). */
  row2?: ReactNode;
  /** Legacy flat children — used by single-row groups that haven't been
   *  migrated to row1/row2 yet. */
  children?: ReactNode;
}) {
  const explicit = row1 !== undefined || row2 !== undefined || lead !== undefined;
  // role="group" + aria-label gives screen readers a way to announce
  // "History group" / "Clipboard group" / "Font group" when the user
  // crosses a boundary, instead of reading every button as a flat list
  // (audit finding 3.1). The visible <div class="ribbon__group-label">
  // already carries the same text for sighted users — the role just
  // exposes it to AT. id="" + aria-labelledby would also work but
  // aria-label is shorter and the label text never differs from
  // what's visible.
  return (
    <div
      className="ribbon__group"
      role="group"
      aria-label={label}
      data-testid={`ribbon-group-${label.toLowerCase()}`}
    >
      {explicit ? (
        <div className="ribbon__group-body ribbon__group-body--two-row">
          {lead && <div className="ribbon__group-lead">{lead}</div>}
          <div className="ribbon__group-stack">
            <div className="ribbon__group-row" data-row="1">{row1}</div>
            <div className="ribbon__group-row" data-row="2">{row2}</div>
          </div>
        </div>
      ) : (
        <div className={`ribbon__group-body${rows ? ' ribbon__group-body--rows' : ''}`}>
          {children}
        </div>
      )}
      <div className="ribbon__group-label" aria-hidden="true">{label}</div>
    </div>
  );
}

/**
 * Big "primary" button — icon stacked on top, label below, optional
 * dropdown chevron. Spans both rows of its RibbonGroup. Use sparingly:
 * one per group at most, reserved for the group's anchor action
 * (Paste, Styles, Format as Table). Matches Excel ribbon visuals.
 */
export function BigToolbarButton({
  id,
  label,
  icon,
  disabled,
  hasChevron,
  onClick,
}: {
  id: string;
  label: string;
  icon: string;
  disabled?: boolean;
  hasChevron?: boolean;
  onClick?: () => void;
}) {
  const split = splitShortcut(label);
  return (
    <Tooltip label={split.label} shortcut={split.shortcut}>
      <button
        type="button"
        className="btn btn--big"
        data-testid={`ribbon-btn-${id}`}
        aria-label={label}
        disabled={disabled}
        onClick={onClick}
      >
        <Icon name={icon} size="lg" />
        <span className="btn__big-label">{split.label}</span>
        {hasChevron && <Icon name="arrow_drop_down" size="sm" className="btn__big-chevron" />}
      </button>
    </Tooltip>
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

/**
 * Split a label like `"Bold (Ctrl+B)"` into `{ label: "Bold", shortcut:
 * "Ctrl+B" }`. Lets toolbar callsites keep using the existing
 * "Name (Shortcut)" string convention while we render the shortcut as
 * a styled pill in the tooltip instead of inline text. Returns the
 * original label and no shortcut when nothing matches.
 */
function splitShortcut(label: string): { label: string; shortcut?: string } {
  // Match a trailing `(...)` containing a Ctrl/Cmd/Alt/Shift/F-key/Enter/Tab
  // sequence — avoids hijacking parens used for prose like "Margins
  // (normal)".
  const m = label.match(/^(.*?)\s*\(((?:Ctrl|Cmd|Alt|Shift|F\d{1,2}|Enter|Tab|Esc)[^)]*)\)\s*$/);
  if (!m) return { label };
  return { label: m[1], shortcut: m[2] };
}

export function ToolbarButton({
  id,
  label,
  icon,
  pressed,
  disabled,
  onClick,
}: ToolbarButtonProps) {
  const split = splitShortcut(label);
  return (
    <Tooltip label={split.label} shortcut={split.shortcut}>
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
 * change it" behavior. Same is true for the line-weight picker.
 */
const BORDER_COLOR_PRESETS = [
  '#000000', '#666666', '#9aa0a6', '#d2d6dc',
  '#d93025', '#e8710a', '#f9ab00', '#188038',
  '#1a73e8', '#7627bb', '#a142f4', '#e91e63',
];

export type BorderWeightChoice = 'thin' | 'medium' | 'thick';

const BORDER_WEIGHT_OPTIONS: Array<{ id: BorderWeightChoice; label: string; px: number }> = [
  { id: 'thin', label: 'Thin', px: 1 },
  { id: 'medium', label: 'Medium', px: 2 },
  { id: 'thick', label: 'Thick', px: 3 },
];

export function BordersControl({
  label,
  icon,
  items,
  disabled,
  defaultColor,
  defaultWeight = 'thin',
  onChoose,
  onDefault,
}: {
  label: string;
  icon: string;
  items: DropdownItem[];
  disabled?: boolean;
  /** Initial color before the user picks one. */
  defaultColor: string;
  /** Initial line weight before the user picks one. */
  defaultWeight?: BorderWeightChoice;
  /** Fires with the chosen border style, current color, AND current weight. */
  onChoose: (itemId: string, color: string, weight: BorderWeightChoice) => void;
  /** Fires when the icon (not the caret) is clicked. Receives the current
   *  color and weight so the icon-click applies "All borders" with the
   *  user's session preferences. */
  onDefault?: (color: string, weight: BorderWeightChoice) => void;
}) {
  const anchorRef = useRef<HTMLDivElement>(null);
  const [open, setOpen] = useState(false);
  const [color, setColor] = useState(defaultColor);
  const [weight, setWeight] = useState<BorderWeightChoice>(defaultWeight);

  return (
    <span
      ref={anchorRef}
      className="btn-split"
      data-testid="ribbon-dropdown-borders"
    >
      <Tooltip label={`${label} (${weight}, ${color})`}>
        <button
          type="button"
          className="btn btn--icon btn-split__icon btn-color__icon"
          data-testid="ribbon-dropdown-borders-apply"
          aria-label={label}
          disabled={disabled}
          onClick={() => {
            onDefault?.(color, weight);
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
                onChoose(item.id, color, weight);
                setOpen(false);
              }}
            >
              <Icon name={item.icon} size="sm" className="menu__item-icon" />
              <span>{item.label}</span>
            </button>
          ))}
          <div className="menu__divider" />
          {/* Line-weight picker — thin/medium/thick. Renders a row of
              swatches at the picked thickness so the user can see what
              they're getting. Sticky for the session (same as colour). */}
          <div className="menu__color-row" data-testid="ribbon-dropdown-borders-weight-row">
            <div className="menu__color-label">Line weight</div>
            <div className="menu__color-swatches" role="group" aria-label="Border line weight">
              {BORDER_WEIGHT_OPTIONS.map((opt) => (
                <button
                  key={opt.id}
                  type="button"
                  className={`menu__weight-swatch${weight === opt.id ? ' menu__weight-swatch--active' : ''}`}
                  data-testid={`ribbon-dropdown-borders-weight-${opt.id}`}
                  aria-label={`Set border weight to ${opt.label}`}
                  aria-pressed={weight === opt.id}
                  title={opt.label}
                  onClick={() => setWeight(opt.id)}
                >
                  <span
                    className="menu__weight-swatch-line"
                    style={{ background: color, height: `${opt.px}px` }}
                    aria-hidden="true"
                  />
                </button>
              ))}
            </div>
          </div>
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
