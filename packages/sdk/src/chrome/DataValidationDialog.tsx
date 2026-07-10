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

/**
 * DataValidationDialog — the SDK chrome's built-in Data Validation modal.
 *
 * The exemplar the other 11 built-in dialogs copy: it reads the active A1
 * selection off the FUniver facade, gathers a validation rule via a small form,
 * and applies it through the `@univerjs/sheets-data-validation` facade — the
 * `univerAPI.newDataValidation()` builder → `.build()` → `FRange.setDataValidation(rule)`
 * pattern (grounded in `sheets-data-validation/lib/types/facade/*.d.ts`). Passing
 * `null` to `setDataValidation` clears the rule, which is how "Remove rule" works.
 *
 * Rule types: list of items · whole number · decimal · date · text length ·
 * checkbox. Each maps to a real builder method:
 *   - list            → `.requireValueInList(values, false, true)`
 *   - whole number    → `.requireNumber*(…, isInteger=true)`
 *   - decimal         → `.requireNumber*(…, isInteger=false)`
 *   - date            → `.requireDate*(Date, …)`
 *   - text length     → `.requireFormulaSatisfied('=LEN(<anchor>) <op> n')`
 *                       (the builder has no first-class text-length method — see
 *                        the facade note below; formula-satisfied is the real,
 *                        installed fallback)
 *   - checkbox        → `.requireCheckbox()`
 *
 * Mounted by `<DialogHost>` when `openDialog('data-validation')` is called and no
 * host override is registered.
 */

import { useMemo, useState, type CSSProperties } from 'react';
// Side-effect import: installs `FUniver.newDataValidation()` and
// `FRange.setDataValidation / getDataValidation` on the facade prototype (both
// the runtime methods AND the TS augmentation of `@univerjs/core/facade` +
// `@univerjs/sheets/facade`). The `dv` plugin group in univer/lazy-plugins.ts
// registers the plugin but does NOT import this facade module, so without it the
// builder/setDataValidation calls below are undefined at runtime. Mirrors the
// `drawing` / `threadComment` groups' `/facade` side-effect imports.
import '@univerjs/sheets-data-validation/facade';
import type { DialogComponentProps } from './extensions';
import type { CasualSheetsAPI } from '../sheets/api';
import { Dialog } from './Dialog';
import {
  DIALOG_BTN_PRIMARY_STYLE,
  DIALOG_BTN_SECONDARY_STYLE,
  DIALOG_FIELD_STYLE,
  DIALOG_INPUT_STYLE,
  DIALOG_LABEL_STYLE,
} from './dialog-styles';

/** Rule families the dialog can build. */
type RuleType = 'list' | 'number' | 'decimal' | 'date' | 'textLength' | 'checkbox';

/** Operators for the numeric / text-length / date families. */
type Operator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greater'
  | 'greaterEqual'
  | 'less'
  | 'lessEqual';

const RULE_TYPE_OPTIONS: Array<{ value: RuleType; label: string }> = [
  { value: 'list', label: 'List of items' },
  { value: 'number', label: 'Whole number' },
  { value: 'decimal', label: 'Decimal' },
  { value: 'date', label: 'Date' },
  { value: 'textLength', label: 'Text length' },
  { value: 'checkbox', label: 'Checkbox' },
];

const OPERATOR_OPTIONS: Array<{ value: Operator; label: string }> = [
  { value: 'between', label: 'Between' },
  { value: 'notBetween', label: 'Not between' },
  { value: 'equal', label: 'Equal to' },
  { value: 'notEqual', label: 'Not equal to' },
  { value: 'greater', label: 'Greater than' },
  { value: 'greaterEqual', label: 'Greater than or equal to' },
  { value: 'less', label: 'Less than' },
  { value: 'lessEqual', label: 'Less than or equal to' },
];

/** True when the operator needs a second operand. */
function isRangeOperator(op: Operator): boolean {
  return op === 'between' || op === 'notBetween';
}

/** Rule types that carry an operator + operand(s). */
function usesOperator(type: RuleType): boolean {
  return type === 'number' || type === 'decimal' || type === 'date' || type === 'textLength';
}

interface DialogState {
  ruleType: RuleType;
  operator: Operator;
  /** Numeric / date / text-length first operand (raw string; parsed on apply). */
  operand1: string;
  /** Second operand, used only for between / notBetween. */
  operand2: string;
  /** Newline-separated list items for the 'list' type. */
  listItems: string;
  ignoreBlank: boolean;
  showErrorMessage: boolean;
  errorMessage: string;
  showInputMessage: boolean;
  inputMessage: string;
}

const INITIAL_STATE: DialogState = {
  ruleType: 'list',
  operator: 'between',
  operand1: '',
  operand2: '',
  listItems: '',
  ignoreBlank: true,
  showErrorMessage: false,
  errorMessage: '',
  showInputMessage: false,
  inputMessage: '',
};

/** The active FRange, or null when there is no selection. */
function activeRange(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange() ?? null;
}

/**
 * Build an FDataValidation rule from the form via the sheets-data-validation
 * builder, then apply it (or clear, when `remove` is true). Returns false when
 * there's no active range or the inputs are unusable.
 */
function applyValidation(api: CasualSheetsAPI, s: DialogState, remove: boolean): boolean {
  const range = activeRange(api);
  if (!range) return false;

  // The FRange setDataValidation extension is contributed by the
  // sheets-data-validation facade module; typed loosely here so the SDK doesn't
  // hard-depend on the facade's ambient FRange augmentation at this call site.
  const dvRange = range as unknown as {
    setDataValidation: (rule: unknown | null) => unknown;
    getA1Notation?: () => string;
  };

  if (remove) {
    dvRange.setDataValidation(null);
    return true;
  }

  const builder = api.univer.newDataValidation();
  const n1 = Number(s.operand1);
  const n2 = Number(s.operand2);
  const wholeNumber = s.ruleType === 'number';

  switch (s.ruleType) {
    case 'list': {
      const values = s.listItems
        .split(/[\n,]/)
        .map((v) => v.trim())
        .filter((v) => v.length > 0);
      if (values.length === 0) return false;
      // (values, multiple=false, showDropdown=true)
      builder.requireValueInList(values, false, true);
      break;
    }
    case 'number':
    case 'decimal': {
      if (!Number.isFinite(n1)) return false;
      switch (s.operator) {
        case 'between':
          if (!Number.isFinite(n2)) return false;
          builder.requireNumberBetween(n1, n2, wholeNumber);
          break;
        case 'notBetween':
          if (!Number.isFinite(n2)) return false;
          builder.requireNumberNotBetween(n1, n2, wholeNumber);
          break;
        case 'equal':
          builder.requireNumberEqualTo(n1, wholeNumber);
          break;
        case 'notEqual':
          builder.requireNumberNotEqualTo(n1, wholeNumber);
          break;
        case 'greater':
          builder.requireNumberGreaterThan(n1, wholeNumber);
          break;
        case 'greaterEqual':
          builder.requireNumberGreaterThanOrEqualTo(n1, wholeNumber);
          break;
        case 'less':
          builder.requireNumberLessThan(n1, wholeNumber);
          break;
        case 'lessEqual':
          builder.requireNumberLessThanOrEqualTo(n1, wholeNumber);
          break;
      }
      break;
    }
    case 'date': {
      const d1 = new Date(s.operand1);
      const d2 = new Date(s.operand2);
      if (Number.isNaN(d1.getTime())) return false;
      switch (s.operator) {
        case 'between':
          if (Number.isNaN(d2.getTime())) return false;
          builder.requireDateBetween(d1, d2);
          break;
        case 'notBetween':
          if (Number.isNaN(d2.getTime())) return false;
          builder.requireDateNotBetween(d1, d2);
          break;
        case 'equal':
          builder.requireDateEqualTo(d1);
          break;
        // The builder has no requireDateNotEqualTo; map the remaining operators
        // to the closest available on/after / on/before methods.
        case 'notEqual':
        case 'greater':
          builder.requireDateAfter(d1);
          break;
        case 'greaterEqual':
          builder.requireDateOnOrAfter(d1);
          break;
        case 'less':
          builder.requireDateBefore(d1);
          break;
        case 'lessEqual':
          builder.requireDateOnOrBefore(d1);
          break;
      }
      break;
    }
    case 'textLength': {
      if (!Number.isFinite(n1)) return false;
      // No first-class text-length rule in the installed builder — express it as
      // a per-cell LEN() formula on the range's top-left anchor. Univer offsets
      // the reference relative to each cell in the range (see requireFormulaSatisfied).
      const anchor = dvRange.getA1Notation?.() ?? 'A1';
      const cell = anchor.split(':')[0];
      const len = `LEN(${cell})`;
      let formula: string;
      switch (s.operator) {
        case 'between':
          if (!Number.isFinite(n2)) return false;
          formula = `=AND(${len}>=${n1}, ${len}<=${n2})`;
          break;
        case 'notBetween':
          if (!Number.isFinite(n2)) return false;
          formula = `=OR(${len}<${n1}, ${len}>${n2})`;
          break;
        case 'equal':
          formula = `=${len}=${n1}`;
          break;
        case 'notEqual':
          formula = `=${len}<>${n1}`;
          break;
        case 'greater':
          formula = `=${len}>${n1}`;
          break;
        case 'greaterEqual':
          formula = `=${len}>=${n1}`;
          break;
        case 'less':
          formula = `=${len}<${n1}`;
          break;
        case 'lessEqual':
          formula = `=${len}<=${n1}`;
          break;
      }
      builder.requireFormulaSatisfied(formula);
      break;
    }
    case 'checkbox':
      builder.requireCheckbox();
      break;
  }

  builder.setAllowBlank(s.ignoreBlank);
  builder.setOptions({
    showErrorMessage: s.showErrorMessage,
    error: s.showErrorMessage ? s.errorMessage : undefined,
    showInputMessage: s.showInputMessage,
    prompt: s.showInputMessage ? s.inputMessage : undefined,
  });

  dvRange.setDataValidation(builder.build());
  return true;
}

const CHECK_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 6,
  marginBottom: 8,
  cursor: 'pointer',
};

const TEXTAREA_STYLE: CSSProperties = {
  ...DIALOG_INPUT_STYLE,
  height: 'auto',
  minHeight: 64,
  padding: '6px 8px',
  resize: 'vertical',
  lineHeight: 1.4,
};

const RANGE_NOTE_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  marginBottom: 12,
};

const TWO_COL_STYLE: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '1fr 1fr',
  gap: 8,
};

export function DataValidationDialog({ api, onClose }: DialogComponentProps) {
  const [state, setState] = useState<DialogState>(INITIAL_STATE);

  // Read the selection once for the header hint. `getA1Notation` off the live
  // FRange (verified in @univerjs/sheets/facade f-range.d.ts) gives the
  // user-facing A1 label, e.g. "A1:B2".
  const rangeLabel = useMemo(() => {
    const fRange = activeRange(api) as unknown as { getA1Notation?: () => string } | null;
    return fRange?.getA1Notation?.() ?? null;
  }, [api]);

  const hasSelection = activeRange(api) !== null;

  const update = <K extends keyof DialogState>(key: K, value: DialogState[K]) =>
    setState((prev) => ({ ...prev, [key]: value }));

  const apply = () => {
    if (applyValidation(api, state, false)) onClose();
  };

  const removeRule = () => {
    applyValidation(api, state, true);
    onClose();
  };

  const showOperator = usesOperator(state.ruleType);
  const showSecondOperand = showOperator && isRangeOperator(state.operator);
  const operandType = state.ruleType === 'date' ? 'date' : 'number';

  return (
    <Dialog
      title="Data validation"
      onClose={onClose}
      width={440}
      data-testid="cs-data-validation-dialog"
      footer={
        <>
          <button
            type="button"
            style={DIALOG_BTN_SECONDARY_STYLE}
            data-testid="cs-data-validation-remove"
            onClick={removeRule}
          >
            Remove rule
          </button>
          <span style={{ flex: 1 }} />
          <button type="button" style={DIALOG_BTN_SECONDARY_STYLE} onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            style={DIALOG_BTN_PRIMARY_STYLE}
            data-testid="cs-data-validation-apply"
            disabled={!hasSelection}
            onClick={apply}
          >
            Apply
          </button>
        </>
      }
    >
      {hasSelection ? (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-data-validation-range">
          Applies to <strong>{rangeLabel ?? 'the current selection'}</strong>
        </div>
      ) : (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-data-validation-no-selection">
          Select one or more cells first, then reopen this dialog.
        </div>
      )}

      <label style={DIALOG_FIELD_STYLE}>
        <span style={DIALOG_LABEL_STYLE}>Criteria</span>
        <select
          style={DIALOG_INPUT_STYLE}
          data-testid="cs-data-validation-type"
          value={state.ruleType}
          onChange={(e) => update('ruleType', e.target.value as RuleType)}
        >
          {RULE_TYPE_OPTIONS.map((opt) => (
            <option key={opt.value} value={opt.value}>
              {opt.label}
            </option>
          ))}
        </select>
      </label>

      {state.ruleType === 'list' && (
        <label style={DIALOG_FIELD_STYLE}>
          <span style={DIALOG_LABEL_STYLE}>Items (one per line, or comma-separated)</span>
          <textarea
            style={TEXTAREA_STYLE}
            data-testid="cs-data-validation-list-items"
            value={state.listItems}
            placeholder={'Yes\nNo\nMaybe'}
            onChange={(e) => update('listItems', e.target.value)}
          />
        </label>
      )}

      {showOperator && (
        <>
          <label style={DIALOG_FIELD_STYLE}>
            <span style={DIALOG_LABEL_STYLE}>Condition</span>
            <select
              style={DIALOG_INPUT_STYLE}
              data-testid="cs-data-validation-operator"
              value={state.operator}
              onChange={(e) => update('operator', e.target.value as Operator)}
            >
              {OPERATOR_OPTIONS.map((opt) => (
                <option key={opt.value} value={opt.value}>
                  {opt.label}
                </option>
              ))}
            </select>
          </label>

          <div style={showSecondOperand ? TWO_COL_STYLE : undefined}>
            <label style={DIALOG_FIELD_STYLE}>
              <span style={DIALOG_LABEL_STYLE}>{showSecondOperand ? 'Minimum' : 'Value'}</span>
              <input
                style={DIALOG_INPUT_STYLE}
                data-testid="cs-data-validation-operand1"
                type={operandType}
                value={state.operand1}
                onChange={(e) => update('operand1', e.target.value)}
              />
            </label>
            {showSecondOperand && (
              <label style={DIALOG_FIELD_STYLE}>
                <span style={DIALOG_LABEL_STYLE}>Maximum</span>
                <input
                  style={DIALOG_INPUT_STYLE}
                  data-testid="cs-data-validation-operand2"
                  type={operandType}
                  value={state.operand2}
                  onChange={(e) => update('operand2', e.target.value)}
                />
              </label>
            )}
          </div>
        </>
      )}

      {state.ruleType === 'checkbox' && (
        <div style={RANGE_NOTE_STYLE}>
          Each cell in the selection becomes a checkbox (checked = TRUE).
        </div>
      )}

      <label style={CHECK_STYLE} data-testid="cs-data-validation-ignore-blank-label">
        <input
          type="checkbox"
          data-testid="cs-data-validation-ignore-blank"
          checked={state.ignoreBlank}
          onChange={(e) => update('ignoreBlank', e.target.checked)}
        />
        <span>Ignore blank cells</span>
      </label>

      <label style={CHECK_STYLE}>
        <input
          type="checkbox"
          data-testid="cs-data-validation-show-input"
          checked={state.showInputMessage}
          onChange={(e) => update('showInputMessage', e.target.checked)}
        />
        <span>Show input message</span>
      </label>
      {state.showInputMessage && (
        <label style={DIALOG_FIELD_STYLE}>
          <span style={DIALOG_LABEL_STYLE}>Input message</span>
          <input
            style={DIALOG_INPUT_STYLE}
            data-testid="cs-data-validation-input-message"
            value={state.inputMessage}
            onChange={(e) => update('inputMessage', e.target.value)}
          />
        </label>
      )}

      <label style={CHECK_STYLE}>
        <input
          type="checkbox"
          data-testid="cs-data-validation-show-error"
          checked={state.showErrorMessage}
          onChange={(e) => update('showErrorMessage', e.target.checked)}
        />
        <span>Show error message on invalid input</span>
      </label>
      {state.showErrorMessage && (
        <label style={DIALOG_FIELD_STYLE}>
          <span style={DIALOG_LABEL_STYLE}>Error message</span>
          <input
            style={DIALOG_INPUT_STYLE}
            data-testid="cs-data-validation-error-message"
            value={state.errorMessage}
            onChange={(e) => update('errorMessage', e.target.value)}
          />
        </label>
      )}
    </Dialog>
  );
}
